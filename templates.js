/*
 * NEF Vorlagen – Inhalte der drei Mail-Vorlagen
 * ------------------------------------------------
 * Nur diese Datei anpassen, wenn sich Text, Empfänger oder Betreff ändern.
 * Nach dem Hochladen gilt die Änderung sofort für alle Benutzer
 * (Outlook ggf. neu starten). Das Manifest muss dafür NICHT neu verteilt werden.
 *
 * Zeilen:
 *   ""                         -> Leerzeile
 *   "Text"                     -> normale Zeile
 *   { t: "Text", fett: true }  -> fett
 *   { t: "Text", klein: true } -> kleiner Hinweistext (10 pt)
 *
 * kennung: ein Satz aus der Vorlage, an dem das Add-in erkennt, dass die
 *          Vorlage schon eingefügt ist (verhindert doppeltes Einfügen).
 */
var NEF_VORLAGEN = {

  dispo: {
    name: "Dispo Anmeldung",
    betreff: "Dispo Anmeldung",
    an: [{ displayName: "Dispo Porsche Basel", emailAddress: "Dispo@porsche-basel.ch" }],
    cc: [{ displayName: "Gerster Roman", emailAddress: "roman.gerster@porsche-basel.ch" }],
    kennung: "Car Connect Code ist in Dispo bis:",
    zeilen: [
      "Kunde:",
      "",
      "Fz./Modell:",
      "",
      "Ausliefertermin inkl. Uhrzeit:",
      "",
      "Versicherung:",
      { t: "(Bei bestimmter Ansprechperson bitte Name und Tel.-Nr.)", klein: true },
      "",
      "Kontrollschild:",
      "",
      "Einlösungsgrund:",
      { t: "(Achtung: bei Neueinlösung bitte immer eine ID/Passkopie an Dispo weiterleiten.)", klein: true },
      "",
      "Fahrzeugpapiere sind in Dispo bis:",
      "",
      "Car Connect Code ist in Dispo bis:",
      "",
      "Besonderes:",
      { t: "(z.B. Bordmappe in anderer Sprache)", klein: true },
      "",
      "Winterräder/Sommerräder Montage:",
      { t: "(Wenn kein Bescheid, werden die Räder bei Leimgruber eingelagert.)", klein: true },
      "",
      "Eintausch:",
      { t: "(Achtung: bei Firmenkunde bitte noch Rechnungsstellung an uns verlangen!)", klein: true },
      "",
      "Zahlungsart:",
      { t: "(Finanzierung bereits bestätigt/Code 178)", klein: true },
      ""
    ]
  },

  ruecknahme: {
    name: "Rücknahme Eintauschfahrzeug",
    betreff: "Rücknahme Eintauschfahrzeug",
    an: [{ displayName: "Dispo Porsche Basel", emailAddress: "Dispo@porsche-basel.ch" }],
    cc: [{ displayName: "Gerster Roman", emailAddress: "roman.gerster@porsche-basel.ch" }],
    kennung: "Der Fahrzeugschlüssel befindet sich im Safe!",
    zeilen: [
      "Kunde:",
      "",
      "Marke/Modell:",
      "",
      "KM-Stand:",
      "",
      "Standort:",
      "",
      "Mit/ohne Winterräder:",
      "",
      "Händlerfahrzeug/Preis:",
      "",
      { t: "Schadennummer Karosserie:", fett: true },
      "",
      { t: "Schadennummer WSS:", fett: true },
      "",
      "Besonderes:",
      "",
      "Der Fahrzeugschlüssel befindet sich im Safe!",
      ""
    ]
  },

  auftrag: {
    name: "Auftragserweiterung",
    betreff: "Auftragserweiterung",
    an: [],
    cc: [],
    kennung: "bestätigen wir Ihnen die Auftragserweiterung",
    zeilen: [
      "Sehr geehrte*r Herr / Frau ______",
      "",
      "Wie soeben besprochen, bestätigen wir Ihnen die Auftragserweiterung für ________ zum vorangeschlagenen Betrag von CHF ______.",
      "",
      "Sehr gerne stehen wir Ihnen für Fragen oder Informationen jederzeit zur Verfügung.",
      "",
      "Wir danken Ihnen für Ihren Auftrag und wünschen Ihnen weiterhin gute Fahrt.",
      "",
      "Freundliche Grüsse",
      ""
    ]
  }
};
