# Excel-zu-XML Automatisierung (persönliches Projekt)

Dieses Python-Skript dient dazu, strukturierte Daten aus einer Excel-Datei automatisch in verschiedene XML-Dateien zu übertragen. Dabei geht es nicht um Standard-Excel-Auswertungen, sondern um ein gezieltes Einfügen von Daten in spezielle XML-Vorlagen, die mehrere definierte Tabellen enthalten – z. B. für Preislisten, Services, Bewertungen, Einkauf und Verkauf.

## Was das Skript macht:

- Liest Informationen aus einer zentralen Excel-Datei ein, die als Steuerzentrale dient (Tabellenblatt **Infos**).
- Hier können z. B. das Startdatum, Enddatum, der gewünschte Dateiname, aktivierter Bereich (Verkauf/Einkauf) und die Unternehmensnummern angegeben werden.
- Auch die Startnummer für neue Service-Einträge (z. B. 4000 oder 5000) kann dort definiert werden.
- Auf Basis dieser Einstellungen wird automatisch ein neuer Dateiname erstellt und die Daten in mehrere **XML-Dateien** übertragen.
- Die Inhalte aus dem Tabellenblatt **Services** (Produktinfos, Preise, Nummern, Texte usw.) werden in die passenden XML-Tabellen eingetragen – je nachdem, ob z. B. Verkauf oder Einkauf aktiv ist.
- Falls ein Preis zu viele Nachkommastellen hat (mehr als zwei), wird dieser automatisch abgerundet. Zusätzlich wird eine Desktop-Benachrichtigung angezeigt, damit man das direkt mitbekommt.

## Weitere Besonderheiten:

- Unternehmensnummern, bei denen in der Excel-Datei „Ja“ steht, werden automatisch erkannt und verarbeitet. Dadurch entstehen dynamisch mehrere Einträge für jede gültige Nummer.
- Die XML-Dateien, in die geschrieben wird, enthalten vorformatierte Tabellen (z. B. `Verkaufsorganisationen`, `Bewertung`, `Services`, `Detaillierte Beschreibungen`, `Positionen`, usw.). Diese werden ab einer bestimmten Zeile (z. B. ab Zeile 8) geleert und neu befüllt.
- Es werden **keine neuen XML-Dateien von Null erstellt**, sondern vorhandene Vorlagen verwendet, die durch das Skript aktualisiert werden.
- Auch der Preislisten-Export aus einer externen Excel-Datei (falls vorhanden) wird angepasst – inklusive Servicenummern, Währung, Einheit und Preis.
- Das komplette Tool wurde in eine `.exe`-Datei umgewandelt, sodass man es **einfach per Doppelklick starten kann**, ohne Python öffnen oder installieren zu müssen. Die Datei kann z. B. auf dem Desktop oder in einem Projektordner liegen und führt beim Start automatisch alle nötigen Schritte aus.
- Beim Umsetzen musste ich mich intensiv mit dem zugrundeliegenden XML-Format der Excel-Vorlagen auseinandersetzen, da schon kleinste Abweichungen oder falsche Strukturelemente dazu führen können, dass die Datei nicht mehr funktioniert oder beschädigt wird. Dadurch war präzises Einfügen an der richtigen Stelle im XML notwendig – was ich programmgesteuert gelöst habe.

## Warum dieses Projekt?

Dieses Tool entstand als privates Lernprojekt, um mit Python eine flexible Automatisierung zu bauen, die **zwischen Excel und XML vermittelt** – besonders bei streng strukturierten Formaten. Es spart viel manuelle Arbeit und reduziert Fehlerquellen beim Kopieren großer Datenmengen.
