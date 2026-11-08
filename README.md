# Word to PDF (Android)

A simple Android app that converts Word documents (.docx) to PDF, fully offline, with no third-party libraries.

**How it works:** pick a .docx file, tap **Convert & save as PDF**, choose where to save, then **Open** or **Share** the PDF.
You can also use "Open with" or "Share" from WhatsApp, Gmail or a file manager to send a .docx straight to the app.

**What it supports:** text, bold, italic, underline, strikethrough, font sizes, colours, headings, alignment (including justified),
spacing, indentation, bullet and numbered lists, tables, PNG/JPG images, line and page breaks, and the document's page size and margins.

**Limits:** old `.doc` files aren't supported (save them as .docx first). Text boxes, headers and footers, footnotes,
custom fonts and EMF/WMF images are skipped. The layout is close to Word's but not pixel-identical.

---

## Get the APK: option A, GitHub (no software needed)

1. Create a new repository on GitHub. It can be private.
2. Upload this whole folder, including the hidden `.github` folder. You can drag and drop it, or use git:
   ```bash
   git init && git add . && git commit -m "Word to PDF app"
   git branch -M main
   git remote add origin https://github.com/<you>/<repo>.git
   git push -u origin main
   ```
3. Open the **Actions** tab. The **Build APK** workflow runs automatically and takes about 3–5 minutes.
4. When it turns green, open the run, scroll to **Artifacts**, and download **WordToPdf-apk** (a zip file).
5. Unzip it, copy `app-debug.apk` to your phone, and tap it to install. You'll need to allow "Install unknown apps" for your file manager or browser.

> If you uploaded through the web page and the workflow doesn't start, go to **Actions → Build APK → Run workflow**.

## Get the APK: option B, Android Studio

1. Open Android Studio, choose **File → Open**, and select this folder.
2. Wait for the Gradle sync to finish.
3. Choose **Build → Build App Bundle(s) / APK(s) → Build APK(s)**.
4. Click **locate** in the popup. The file is `app/build/outputs/apk/debug/app-debug.apk`.

Or from a terminal with the Android SDK installed: `./gradlew assembleDebug`

## Project layout

```
app/src/main/java/com/wordtopdf/app/
├── MainActivity.java        UI: pick file → convert → open/share
├── docx/DocxParser.java     reads .docx (zip + XML) into a simple model
├── docx/DocxModel.java      paragraphs, runs, tables, images
└── pdf/PdfWriter.java       lays the model out on pages with Android's PdfDocument
```

Min Android version: 7.0 (API 24) · Target: Android 15 (API 35)

The release APK is signed with the debug key so it's installable for testing. Use your own keystore before publishing to Google Play.
