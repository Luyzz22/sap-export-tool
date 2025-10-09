markdown# 🏭 SAP Lieferanten-Export & Reporting Tool

Automatisiertes Python-Tool für Export, Bereinigung und Reporting von SAP-Lieferantendaten mit Excel-Output.

![Python](https://img.shields.io/badge/Python-3.8%2B-blue)
![DSGVO](https://img.shields.io/badge/DSGVO-Konform-green)
![Status](https://img.shields.io/badge/Status-Production%20Ready-success)

## 🎯 Problem & Lösung

**Problem:** Zollabteilungen verbringen täglich 2-3 Stunden mit manuellem SAP-Export und Excel-Aufbereitung.

**Lösung:** Vollautomatische Verarbeitung mit Datenbereinigung, Duplikaterkennung und fertigem Excel-Report in Sekunden.

## ✨ Features

- ✅ **Automatischer SAP-Datenexport** (CSV-basiert, SAP-RFC-ready)
- ✅ **Intelligente Datenbereinigung**:
  - Duplikaterkennung nach Lieferantennummer
  - Telefonnummern-Formatierung (+49-Format)
  - Vollständigkeitsprüfung (Pflichtfelder)
- ✅ **Excel-Report mit 3 Sheets**:
  - Sheet 1: Alle Lieferanten (bereinigte Daten)
  - Sheet 2: Nach Land (Pivot-Tabelle mit Statistik)
  - Sheet 3: Prüfliste (unvollständige Datensätze, rot markiert)
- ✅ **Professionelles Logging**: Detaillierte Log-Datei mit Timestamp und Statistiken
- ✅ **E-Mail-Versand**: Simuliert, produktionsbereit erweiterbar
- ✅ **DSGVO-konform**: 100% lokale Verarbeitung, keine Cloud

## 🚀 Schnellstart

### Installation
```bash
# Repository klonen
git clone https://github.com/luisschenk/sap-export-tool.git
cd sap-export-tool

# Abhängigkeiten installieren
pip3 install -r requirements.txt
Ausführung
bash# Test-Daten erstellen (50 Beispiel-Lieferanten)
python3 create_test_data.py

# Tool ausführen
python3 sap_export_tool.py

# Excel-Report öffnen
open Lieferanten_Report_*.xlsx
📊 Output-Beispiel
Nach Ausführung erhalten Sie:
Excel-Report (Lieferanten_Report_YYYYMMDD_HHMMSS.xlsx):

📄 Sheet 1 "Alle Lieferanten": Vollständige bereinigte Datensätze
📊 Sheet 2 "Nach Land": Pivot-Tabelle (Anzahl Lieferanten pro Land)
⚠️ Sheet 3 "Prüfliste": Unvollständige Datensätze (rot markiert)

Log-Datei (sap_export_log.txt):
[2025-10-09 14:30:15] === NEUER EXPORT GESTARTET ===
[2025-10-09 14:30:15] Eingelesen: 52 Datensätze
[2025-10-09 14:30:15] Duplikate entfernt: 2
[2025-10-09 14:30:15] Unvollständig: 8
[2025-10-09 14:30:16] === EXPORT BEENDET ===
🏢 Produktiv-Einsatz
SAP-Integration (Echtdaten)
Option 1: CSV-Export aus SAP
python# Transaktion SE16 → Tabelle LFA1 → Export als CSV
# Umbenennen zu: sap_lieferanten.csv
Option 2: SAP RFC (mit pyrfc)
pythonfrom pyrfc import Connection

conn = Connection(
    user='SAPUSER', passwd='PASSWORD',
    ashost='sap.server.de', sysnr='00', client='100'
)
result = conn.call('RFC_READ_TABLE', QUERY_TABLE='LFA1')
# Weiterverarbeitung im Tool
Automatisierung (täglich um 2 Uhr)
Windows Task Scheduler:
batchschtasks /create /tn "SAP_Export" /tr "python3 C:\path\sap_export_tool.py" /sc daily /st 02:00
Linux/Mac Cron:
bashcrontab -e
# Zeile hinzufügen:
0 2 * * * cd /path/to/tool && python3 sap_export_tool.py
E-Mail-Versand (Echtbetrieb)
Code in sap_export_tool.py erweitern:
pythonimport smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

def sende_email(datei):
    msg = MIMEMultipart()
    msg['From'] = 'sap-export@firma.de'
    msg['To'] = EMAIL_EMPFAENGER
    msg['Subject'] = f"SAP Lieferanten-Report {datetime.now().strftime('%d.%m.%Y')}"
    
    with open(datei, 'rb') as f:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(f.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename={os.path.basename(datei)}')
        msg.attach(part)
    
    server = smtplib.SMTP('smtp.firma.de', 587)
    server.starttls()
    server.login('username', 'password')
    server.send_message(msg)
    server.quit()
🔧 Anpassungen
Pflichtfelder ändern
In sap_export_tool.py, Zeile 12:
pythonPFLICHTFELDER = ['Lieferantennr', 'Name', 'Land', 'Strasse']  # Nach Bedarf erweitern
Weitere CSV-Spalten hinzufügen
In create_test_data.py:
pythondatensatz = {
    'Lieferantennr': ...,
    'Name': ...,
    'Umsatz': random.randint(10000, 500000),  # Neu
    'Kategorie': random.choice(['A', 'B', 'C'])  # Neu
}
🔐 DSGVO & Datenschutz

✅ 100% lokale Verarbeitung: Keine Cloud-Uploads, keine externen APIs
✅ Firmen-Stammdaten: Keine personenbezogenen Daten (DSGVO Art. 4)
✅ Logging ohne sensible Daten: Nur Statistiken, keine Namen/Adressen
✅ Verschlüsselung optional: Excel-Passwort-Schutz erweiterbar

📈 Performance

50 Datensätze: < 1 Sekunde
1.000 Datensätze: ~2 Sekunden
10.000 Datensätze: ~15 Sekunden

Getestet auf: MacBook Air M1, Python 3.13
🛠️ Technologie-Stack

Python 3.8+
pandas: Excel-Verarbeitung, Datenanalyse
openpyxl: Excel-Formatierung (Farben, Styles)
datetime: Timestamps, Logging
re: Telefonnummern-Formatierung (Regex)

🧪 Testing
bash# Test-Daten mit verschiedenen Szenarien
python3 create_test_data.py

# Erwartetes Ergebnis:
# - 50 Lieferanten generiert
# - ~15% unvollständig (für Prüfliste)
# - 2 Duplikate (zum Testen der Erkennung)
📄 Use Cases

Zollabteilungen: Täglicher Lieferanten-Export für Compliance-Checks
Einkauf: Lieferanten-Stammdaten für Verhandlungen aufbereiten
Controlling: Lieferanten-Statistiken nach Land/Region
Qualitätsmanagement: Unvollständige Datensätze identifizieren

🤝 Beitragen
Feedback und Verbesserungsvorschläge willkommen!
Kontakt: Luis@schenk.com

👤 Autor
Luis Schenk
Wirtschaftsjurist | Customs & Trade Compliance Specialist
📧 Luis@schenk.com
📍 Weinheim, Rhein-Neckar-Region, Deutschland
💼 Verfügbar: 15-20h/Woche (Werkstudent), ab 2025 Vollzeit
🎯 Weitere Projekte

Zoll-Präferenzkalkulator - Automatisierte EU-Präferenzursprungsberechnung
Custom GPT: Zoll-Assistent - KI-Assistent für Zollfragen
Custom GPT: Legal Translator - Mehrsprachige Legal-Tech

💼 Services

KI-gestützte Zoll-Automatisierung: Präferenzkalkulationen, Lieferantenerklärungen, Sanktionslistenscreening
SAP/Excel-Integration: Python-basierte Workflow-Automatisierung
Datenbereinigung & Business Intelligence: DSGVO-konforme Lösungen
Custom-Tool-Entwicklung: Maßgeschneiderte Automatisierungen für Compliance & Zoll

📄 Lizenz
MIT License - Für kommerzielle und private Nutzung frei verwendbar.
Version: 1.0.0
Letzte Aktualisierung: Oktober 2025
GitHub: https://github.com/luisschenk/sap-export-tool
