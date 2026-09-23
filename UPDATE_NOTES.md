# DCT Workbench 4.0 — Update Notes

## New in 4.0

- Added a Workbench that brings source validation, the PACS mapping suite, and post-generation validation into one guided workflow.
- Added adjustable thresholds for column-wide negative-sign detection, date checks, installment amounts, deposit maturity amounts, and RD paid totals. Existing defaults are preserved.
- Added local run history with timestamps, input/output file fingerprints, active settings, and a reviewed-report record. It does not store spreadsheet row values.
- Added launch/setup and standalone Windows build scripts.
- Added a phone-first Android companion for selecting multiple template files, validating them on-device, reviewing the summary, and sharing report ZIPs.
- Added a GitHub Actions Android build that publishes a debug APK as a workflow artifact.
- Updated the PACS suite version to 4.0; its mapping and cleanup behavior is retained from the 09-09-2026 v3.5 release.

## Validator 3.1

- Reads `dct_validator_settings.json` and applies configured thresholds to the existing date, amount, and sign convention checks.
- Adds active threshold values to each file summary so a report shows which tolerances were used.
- Keeps the existing validation rules, output shape, cross-file checks, and read-only source handling.

## Review process

The user reviews the generated Rejected Rows and Review Items sheets, then records the review in Workbench. Existing suite cleanup/mapping audit sheets retain details about its transformations. No new silent correction behavior was added.

## Important limits

- Workbench validation runs are recorded automatically. PACS suite launch is recorded, but the suite's completed conversion is not automatically imported into Workbench history; use the suite's own Reports tab and output folder for its audit trail.
- The thresholds are adjustable, but changing a threshold should be based on the institution's confirmed portal behavior.
- A successful local build depends on the installed Python and PyInstaller versions and Windows environment.
- The Android companion covers template validation and report sharing; the full PACS mapping UI remains desktop-only.
