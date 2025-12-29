👉 Dieses Skript automatisiert große Teile des TEMU Uploads.
Es nimmt eine CSV-Datei aus PlentyMarkets, verarbeitet sie und trägt alle benötigten Daten automatisch in eine bestehende TEMU-Vorlage (Excel) ein.

🔄 1. Automatisierter Start & Benutzerhinweise
Beim Start zeigt das Skript:
- eine kurze Einführung
- wichtige Hinweise (z. B. dass die Dateien geschlossen sein müssen)
- einen Schritt-für-Schritt-Workflow für den gesamten Prozess
Damit ist es auch für weniger technische Anwender leicht verständlich.

📥 2. Einlesen der CSV-Datei
Das Skript lädt automatisch:
- die exportierte CSV aus Plenty mit dem korrekten Trennzeichen inklusive Spaltenname-Überprüfung
- Es gibt dem Nutzer direkt Feedback, welche Spalten gefunden wurden.

📐 3. Automatische Umrechnung von Maßen (mm → cm)
Viele TEMU-Felder erwarten Maße in cm.
Plenty liefert diese aber häufig in Millimetern.
Das Skript rechnet deshalb automatisch um:
- Länge
- Breite
- Höhe
Kein manueller Aufwand mehr – die richtigen Einheiten sind garantiert.

📦 4. Ermittlung der Gesamtartikelanzahl
Viele Artikel bestehen aus mehreren Einheiten, die in Plenty oft so gespeichert sind:
12345:1;98765:2;54321:1

Das Skript:
- erkennt dieses Format,
- extrahiert die Stückzahlen,
- berechnet vollautomatisch die Gesamtartikelanzahl.
Bei Fehlern wird eine sinnvolle Mindestmenge eingesetzt, sodass die Daten immer vollständig bleiben.

📊 5. Laden der TEMU-Excel-Vorlage
Das Skript öffnet die bestehende TEMU.xlsx, prüft das richtige Tabellenblatt und bereitet das Eintragen vor.
Fehler wie "falsches Sheet" werden frühzeitig gemeldet.

🗂 6. Intelligentes Spaltenmapping
Das Herzstück des Skripts:
Ein umfangreiches Mapping legt fest, welche CSV-Information in welche Excel-Spalte geschrieben werden soll.
Beispiele:
- SKU → B
- Produktname → K
- Beschreibung → S
- Maße → EE / EF / EG
- Bilder-URLs → Z oder DN

🧹 7. Automatisches Leeren alter Daten
Bevor neue Werte eingetragen werden, löscht das Skript alte Einträge ab einer bestimmten Zeile (standardmäßig ab Zeile 5).
Damit bleiben:
- Kopfzeile,
- Formatierungen,
- Formeln
im Excel erhalten.

✍️ 8. Eintrag der neuen Produktdaten
Zeile für Zeile werden nun:
alle gemappten Felder aus der CSV in die passenden Zellen der Excel-Vorlage geschrieben.
Das ist vollständig automatisiert und ersetzt Stunden manueller Arbeit.

💾 9. Speichern & Abschlussmeldung
Am Ende:
- speichert das Skript die Excel-Datei
- gibt eine Bestätigung aus
- zeigt den Datei-Namen an
- wartet, bis der Anwender den Vorgang abschließt
Damit ist der Prozess klar abgeschlossen.

👉 Vorteile für den Arbeitsalltag
Enorme Zeitersparnis
Statt hunderte Produkte manuell zu pflegen:
Ein Klick → Fertige Importdatei.

Weniger Fehler
- einheitliche Maße
- konsistente Zuordnung
- zuverlässige Berechnung von Mengen
- keine Copy-&-Paste-Fehler mehr

Skalierbar & sicher
Ideal auch für große Datenmengen.
Kann problemlos erweitert werden, z. B.:
- automatische Kategorie-Zuordnung
- zusätzliche Qualitätsprüfungen
- Log-Dateien

👉 Zusammenfassung
Dieses Skript ist ein effektives Automatisierungswerkzeug, das:
- CSV-Daten automatisch verarbeitet,
- Maße umrechnet,
- Mengen berechnet,
- Altdaten löscht,
- die TEMU-Vorlage korrekt befüllt
und so den gesamten Produkt-Upload stark vereinfacht.

👉 Einfach ausführen – und die fertige Excel-Datei importieren.

Updates:
- Ursprungsland/-region -> automatische Übersetzung für TEMU
- Leere Werte in 'Nicht verfügbar für Listenpreis' werden zu 'N/A'
- CSV vollständig als Strings einlesen
- (ausgesetzt) saubere Trennung von Bild URLs
- entfernen von rich text
- aktivieren von erforderlichen nicht-Pflichtfeldern
- 512 Zeichenlimit für "URL für SKU-Bilder"
- Standardwert für 'Anzahl' = 0
- Zeichenkürzung für SKU Bilder und Aufzählungen
  
:: V4.1e ::
- Seperator Bilder
- Seperator Aufzählungspunkte
- automatische Kategorieerkennung
- ignoriere Einträge ohne Produkt ID
- benötigte Spalten ergänzt
- Gesamtartikel = Artikel
- kleinere Fixes

