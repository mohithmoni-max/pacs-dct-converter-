# PACS DCT Workbench 4.0

This release brings the PACS DCT + Mapping Suite and the all-template DCT validator into one guided desktop workflow.

## What is included

- **DCT Workbench:** choose a source folder, run validation, open the PACS suite, then validate the suite's generated upload folder.
- **Configurable validation tolerances:** edit date, amount, and sign-convention thresholds in the Workbench. Defaults match validator v3.0.
- **Local audit history:** record validation runs with file names, SHA-256 fingerprints, settings, report fingerprints, and an explicit report-reviewed timestamp. Spreadsheet row contents are not copied into the history.
- **Review record:** after inspecting the Rejected Rows and Review Items sheets, use “Mark latest report reviewed” to record that review.
- **Simpler launch:** `Launch_DCT_Workbench.bat` installs the listed Python packages and starts the Workbench.
- **Windows app build:** `Build_Windows_Apps.bat` creates standalone applications in `dist` with PyInstaller.
- **Android companion:** the `android/` project provides a phone-sized DCT validator and a GitHub Actions build for a debug APK.

The original PACS suite retains its mapping screens and detailed cleanup/audit reports. The validator continues to write clean data, rejected rows, review items, and consolidated summaries into a dated `Validated_*` folder. Source files are opened read-only by the validator.

## Run from source

1. Install Python 3.10 or newer.
2. Double-click `Launch_DCT_Workbench.bat`.
3. Choose a folder and validate its DCT files.
4. Open the PACS DCT suite to map and generate output.
5. Return to the Workbench and choose “Validate generated output…” before upload.

Alternatively, install `requirements.txt` and run `python app/dct_workbench.py`.

## Settings and local records

`app/dct_validator_settings.json` stores the active rule tolerances. The Workbench lets you change these values and applies them on the next validation. `app/run_history.json` is created after a run and records metadata and hashes locally. Both files can be backed up with the application folder.

Review approval records that a person reviewed a report; it does not change spreadsheet contents or assert that every flagged item is correct. Corrections in the PACS suite remain visible through its existing mapping and cleanup reports.

The Android companion validates and shares reports. The desktop PACS suite's mapping screens remain a Windows workflow.

## Version notes

- PACS DCT + Mapping Suite: 4.0, based on the 09-09-2026 v3.5 source.
- All-Template Validator: 3.1, based on the 17-09-2026 v3.0 source.
- Full update summary: [UPDATE_NOTES.md](UPDATE_NOTES.md).
