import pandas as pd
import random

laender = ['Deutschland', 'Frankreich', 'Italien', 'Spanien', 'Polen']
firmen = ['Mueller GmbH', 'Schmidt AG', 'Weber & Co', 'Fischer KG', 'Meyer Industries']

def generiere_testdaten(anzahl=50):
    daten = []
    for i in range(anzahl):
        ist_unvollstaendig = random.random() < 0.15
        datensatz = {
            'Lieferantennr': f"LF{1000 + i:05d}",
            'Name': random.choice(firmen) if not ist_unvollstaendig else '',
            'Strasse': f"Hauptstrasse {random.randint(1, 150)}",
            'PLZ': f"{random.randint(10000, 99999)}",
            'Ort': random.choice(['Muenchen', 'Berlin', 'Hamburg']),
            'Land': random.choice(laender) if not ist_unvollstaendig else '',
            'Telefon': f"+49{random.randint(100,999)}{random.randint(1000000,9999999)}"
        }
        daten.append(datensatz)
    
    daten.append(daten[5].copy())
    daten.append(daten[12].copy())
    return daten

if __name__ == "__main__":
    print("Erstelle Test-Daten...")
    testdaten = generiere_testdaten(50)
    df = pd.DataFrame(testdaten)
    df.to_csv("sap_lieferanten.csv", index=False, encoding='utf-8')
    print("sap_lieferanten.csv erstellt!")
    print("Starte nun: python3 sap_export_tool.py")
