/*
 * NEF Vorlagen – Outlook-Add-in (Funktionsbefehle ohne Aufgabenbereich)
 *
 * Zwei Arten von Buttons pro Vorlage:
 *   ...Einfuegen : im Fenster "Neue E-Mail". Ergänzt An/CC, setzt den Betreff
 *                  (falls leer) und fügt den Text OBERHALB der Signatur ein.
 *   ...Neu       : im Hauptfenster (bei markierter Mail). Öffnet direkt eine
 *                  neue, vorausgefüllte Mail.
 *
 * Inhalte stehen in templates.js (NEF_VORLAGEN).
 */
/* global Office, NEF_VORLAGEN */

(function () {
  "use strict";

  var FEHLER_ID = "nefVorlageFehler";

  // ---------- HTML / Text aus den Vorlagenzeilen erzeugen ----------

  function esc(s) {
    return String(s)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }

  function zeile(z) {
    return typeof z === "string" ? { t: z } : z;
  }

  // Bewusst ohne Inline-CSS (laut Microsoft von prependAsync/setAsync nicht
  // unterstützt) – nur einfache Tags, die alle Outlook-Editoren verstehen.
  function zeileHtml(z) {
    z = zeile(z);
    if (!z.t) return "<div><br></div>";
    var inhalt = esc(z.t);
    if (z.fett) inhalt = "<b>" + inhalt + "</b>";
    var font = z.klein ? '<font face="Arial" size="2">' : '<font face="Arial">';
    return "<div>" + font + inhalt + "</font></div>";
  }

  function vorlageHtml(v) {
    return v.zeilen.map(zeileHtml).join("");
  }

  function vorlageText(v) {
    return v.zeilen.map(function (z) { return zeile(z).t || ""; }).join("\n") + "\n";
  }

  // ---------- Promise-Hülle für die Office-Callback-API ----------

  function aufruf(obj, methode) {
    var args = Array.prototype.slice.call(arguments, 2);
    return new Promise(function (resolve, reject) {
      args.push(function (r) {
        if (r.status === Office.AsyncResultStatus.Succeeded) resolve(r.value);
        else reject(r.error || new Error("Unbekannter Fehler"));
      });
      obj[methode].apply(obj, args);
    });
  }

  function fehlerAnzeigen(item, text) {
    try {
      if (item && item.notificationMessages) {
        item.notificationMessages.replaceAsync(FEHLER_ID, {
          type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
          message: String(text).slice(0, 150)
        });
      }
    } catch (e) { /* nichts mehr zu tun */ }
  }

  // Empfänger nur ergänzen, wenn sie noch nicht drin sind.
  function empfaengerErgaenzen(feld, liste) {
    if (!liste || !liste.length) return Promise.resolve();
    return aufruf(feld, "getAsync").then(function (vorhanden) {
      var bekannt = (vorhanden || []).map(function (e) {
        return String(e.emailAddress || "").toLowerCase();
      });
      var neu = liste.filter(function (e) {
        return bekannt.indexOf(e.emailAddress.toLowerCase()) === -1;
      });
      return neu.length ? aufruf(feld, "addAsync", neu) : null;
    });
  }

  // ---------- Button im Fenster "Neue E-Mail" ----------

  function inMailEinfuegen(event, schluessel) {
    var v = NEF_VORLAGEN[schluessel];
    var item = Office.context.mailbox.item;

    var ablauf = empfaengerErgaenzen(item.to, v.an)
      .then(function () { return empfaengerErgaenzen(item.cc, v.cc); })
      .then(function () { return aufruf(item.subject, "getAsync"); })
      .then(function (betreff) {
        if (!betreff || !String(betreff).trim()) {
          return aufruf(item.subject, "setAsync", v.betreff);
        }
      })
      .then(function () { return aufruf(item.body, "getAsync", Office.CoercionType.Text); })
      .then(function (text) {
        if (v.kennung && String(text || "").indexOf(v.kennung) !== -1) {
          return; // Vorlage steht schon in der Mail – nicht doppelt einfügen
        }
        return aufruf(item.body, "getTypeAsync").then(function (typ) {
          if (typ === Office.CoercionType.Html) {
            return aufruf(item.body, "prependAsync", vorlageHtml(v), { coercionType: Office.CoercionType.Html });
          }
          return aufruf(item.body, "prependAsync", vorlageText(v), { coercionType: Office.CoercionType.Text });
        });
      })
      .then(function () {
        try { item.notificationMessages.removeAsync(FEHLER_ID); } catch (e) { /* egal */ }
      })
      .catch(function (err) {
        fehlerAnzeigen(item, "Vorlage «" + v.name + "» konnte nicht eingefügt werden: " + (err && err.message ? err.message : err));
      });

    ablauf.then(function () { event.completed(); }, function () { event.completed(); });
  }

  // ---------- Button im Hauptfenster: neue Mail öffnen ----------

  function neueMailOeffnen(event, schluessel) {
    var v = NEF_VORLAGEN[schluessel];
    try {
      Office.context.mailbox.displayNewMessageForm({
        toRecipients: v.an.map(function (e) { return e.emailAddress; }),
        ccRecipients: v.cc.map(function (e) { return e.emailAddress; }),
        subject: v.betreff,
        htmlBody: vorlageHtml(v)
      });
    } catch (err) {
      fehlerAnzeigen(Office.context.mailbox.item, "Neue Mail «" + v.name + "» konnte nicht geöffnet werden: " + (err && err.message ? err.message : err));
    }
    event.completed();
  }

  // ---------- Registrierung ----------
  // Global (für das XML-Manifest) und über Office.actions (neuere Clients).

  var befehle = {
    nefDispoEinfuegen:      function (e) { inMailEinfuegen(e, "dispo"); },
    nefDispoNeu:            function (e) { neueMailOeffnen(e, "dispo"); },
    nefRuecknahmeEinfuegen: function (e) { inMailEinfuegen(e, "ruecknahme"); },
    nefRuecknahmeNeu:       function (e) { neueMailOeffnen(e, "ruecknahme"); },
    nefAuftragEinfuegen:    function (e) { inMailEinfuegen(e, "auftrag"); },
    nefAuftragNeu:          function (e) { neueMailOeffnen(e, "auftrag"); }
  };

  var g = typeof window !== "undefined" ? window : this;
  Object.keys(befehle).forEach(function (name) { g[name] = befehle[name]; });

  // Für den lokalen Test (test/test.html) zugänglich machen
  g.NEF_INTERN = { vorlageHtml: vorlageHtml, vorlageText: vorlageText };

  Office.onReady(function () {
    try {
      if (Office.actions && Office.actions.associate) {
        Object.keys(befehle).forEach(function (name) { Office.actions.associate(name, befehle[name]); });
      }
    } catch (e) { /* ältere Clients: globale Funktionen reichen */ }
  });
})();
