# Supplier CSV reporting

A small Python utility that cleans an existing supplier CSV export and writes an Excel review workbook. It has no direct SAP connection.

## What it does

- Keeps the first row for each `Lieferantennr`.
- Normalizes telephone strings with a German `+49` default.
- Checks whether supplier number, name and country are present.
- Writes `Alle Lieferanten`, `Nach Land` and `Pruefliste` worksheets.

## Run with synthetic data

```bash
git clone https://github.com/Luyzz22/sap-export-tool.git
cd sap-export-tool
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
python create_test_data.py
python sap_export_tool.py
```

On Windows, activate with `.venv\Scripts\activate`. The generator overwrites `sap_lieferanten.csv`; run it in a separate working directory when preserving your own inputs.

The input columns used by the program are `Lieferantennr`, `Name`, `Land` and `Telefon`. Additional columns are carried into the workbook. Reports use the name `Lieferanten_Report_<timestamp>.xlsx`; the run log is `sap_export_log.txt`.

## Verify

```bash
python -m unittest discover -s tests -v
```

The test runs the actual command in a temporary directory with synthetic rows. It checks duplicate removal, the missing-field review sheet and workbook structure. GitHub Actions repeats this check with Python 3.12.

## Boundaries

This is a CSV-processing utility, not a validated SAP integration or a compliance product. Telephone normalization does not validate international phone numbers. Missing columns and malformed data need caller-side validation. No mail is sent, and the code does not implement RFC, scheduling or cloud deployment.

Use synthetic data for examples and reviews. Local processing alone does not establish data-protection compliance. Generated inputs, reports and logs are excluded from new commits; removing a tracked sample does not erase earlier Git history.

Maintained by [Luis Schenk](https://github.com/Luyzz22). For current SBS products, see [SBS Deutschland](https://sbsdeutschland.com/).
