# Release 1.1

## macOS fix

- Fix writing generated files when macOS Gatekeeper launches the app from an App Translocation path.
- Embed the vocabulary database and macro-enabled Excel template in the app bundle's Contents/Resources directory.
- Verify the app still generates both regular and macro-enabled Excel workbooks after the app bundle is relocated.
