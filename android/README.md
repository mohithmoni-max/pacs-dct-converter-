# DCT Mobile Validator

Android companion for the DCT all-template validator. The phone-first layout uses full-width touch controls and scrollable content. It validates selected DCT spreadsheets on the device, creates the same detailed Excel reports, and shares the report bundle as a ZIP. Input files are copied into the app's private storage for the run; no network permission or upload service is used by the app.

## Build

Buildozer runs on Linux or macOS; on Windows, use WSL. The included GitHub Actions workflow builds a debug APK and publishes it as the `dct-mobile-validator-apk` workflow artifact. The build uses the python-for-android `develop` branch for the current pandas and NumPy recipes.

To build manually from this folder:

```sh
python -m pip install --upgrade pip setuptools wheel "Cython<4" buildozer==1.6.0
buildozer android debug
```

The first Android build downloads the SDK, NDK, and Python package sources and can take a while. The output APK is written to `bin/`.

## Use

1. Install the debug APK on an Android device.
2. Choose one or more DCT template files with the Android file picker.
3. Run validation and review the summary.
4. Share the ZIP report bundle to Files, Drive, or another destination.
5. Use “Mark report reviewed” after inspecting the report.

The app does not include the Windows PACS mapping screens. Use the desktop PACS DCT suite for source conversion and product mapping, then use the Android app for validation and report review.
