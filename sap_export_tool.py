import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import os
import sys
from datetime import datetime
import re

INPUT_CSV = "sap_lieferanten.csv"
OUTPUT_EXCEL = "Lieferanten_Report_{}.xlsx".format(datetime.now().strftime("%Y%m%d_%H%M%S"))
LOG_FILE = "sap_export_log.txt"
EMAIL_EMPFAENGER = "zoll-team@kunde.de"
PFLICHTFELDER = ['Lieferantennr', 'Name', 'Land']

def log_message(message):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    log_entry = f"[{timestamp}] {message}"
    print(log_entry)
    with open(LOG_FILE, 'a', encoding='utf-8') as f:
        f.write(log_entry + "\n")

def bereinige_telefonnummer(telefon):
    if pd.isna(telefon) or telefon == "":
        return ""
    telefon_clean = re.sub(r'[^\d+]', '', str(telefon))
    if not telefon_clean.startswith('+'):
        if telefon_clean.startswith('0'):
            telefon_clean = '+49' + telefon_clean[1:]
        else:
            telefon_clean = '+49' + telefon_clean
    return telefon_clean

def pruefe_vollstaendigkeit(row):
    for feld in PFLICHTFELDER:
        if pd.isna(row[feld]) or row[feld] == "":
            return False
    return True

def hauptprogramm():
    print("=" * 70)
    print("SAP LIEFERANTEN-EXPORT & REPORTING TOOL")
    print("=" * 70)
    log_message("=== NEUER EXPORT GESTARTET ===")
    
    if not os.path.exists(INPUT_CSV):
        print(f"FEHLER: {INPUT_CSV} nicht gefunden!")
        print("Fuehre erst: python3 create_test_data.py")
        sys.exit(1)
    
    df = pd.read_csv(INPUT_CSV, encoding='utf-8')
    log_message(f"Eingelesen: {len(df)} Datensaetze")
    
    anzahl_vorher = len(df)
    duplikate_vorher = len(df)
    df = df.drop_duplicates(subset=['Lieferantennr'], keep='first')
    duplikate_entfernt = duplikate_vorher - len(df)
    
    df['Telefon'] = df['Telefon'].apply(bereinige_telefonnummer)
    df['Vollstaendig'] = df.apply(pruefe_vollstaendigkeit, axis=1)
    anzahl_unvollstaendig = len(df[df['Vollstaendig'] == False])
    
    log_message(f"Duplikate entfernt: {duplikate_entfernt}")
    log_message(f"Unvollstaendig: {anzahl_unvollstaendig}")
    
    with pd.ExcelWriter(OUTPUT_EXCEL, engine='openpyxl') as writer:
        df_output = df.drop('Vollstaendig', axis=1)
        df_output.to_excel(writer, sheet_name='Alle Lieferanten', index=False)
        
        pivot = df.groupby('Land').size().reset_index(name='Anzahl')
        pivot = pivot.sort_values('Anzahl', ascending=False)
        pivot.to_excel(writer, sheet_name='Nach Land', index=False)
        
        df_unvollstaendig = df[df['Vollstaendig'] == False].drop('Vollstaendig', axis=1)
        df_unvollstaendig.to_excel(writer, sheet_name='Pruefliste', index=False)
    
    wb = load_workbook(OUTPUT_EXCEL)
    if 'Pruefliste' in wb.sheetnames:
        ws = wb['Pruefliste']
        rot = PatternFill(start_color="FFB6C1", end_color="FFB6C1", fill_type="solid")
        for zeile in range(2, ws.max_row + 1):
            for spalte in range(1, ws.max_column + 1):
                ws.cell(row=zeile, column=spalte).fill = rot
    wb.save(OUTPUT_EXCEL)
    
    print("\n" + "=" * 70)
    print("ERFOLGREICH ABGESCHLOSSEN!")
    print(f"Excel-Report: {OUTPUT_EXCEL}")
    print(f"Log-Datei: {LOG_FILE}")
    print("=" * 70)
    log_message("=== EXPORT BEENDET ===")

if __name__ == "__main__":
    hauptprogramm()
