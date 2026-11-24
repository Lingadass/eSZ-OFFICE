# Publishing eSZ OFFICE on Google Play Store (Trusted Web Activity / TWA)

eSZ OFFICE is engineered as an offline-first Progressive Web App suitable for packaging into an Android App Bundle (`.aab`) using **Bubblewrap** or **PWABuilder**.

---

## 1. Pre-Packaging Requirements

1. **HTTPS Hosting**: Deploy the production `dist/` output to an HTTPS domain with valid TLS certificates (e.g., `https://app.eszoffice.com`).
2. **HTTP Security Headers**: Ensure your hosting provider / reverse proxy sends:
   - `Content-Security-Policy`: `default-src 'self'; script-src 'self' 'wasm-unsafe-eval'; style-src 'self' 'unsafe-inline' https://fonts.googleapis.com; font-src 'self' https://fonts.gstatic.com data:; img-src 'self' data: blob:; worker-src 'self' blob:; connect-src 'self'; object-src 'none'; base-uri 'self'; form-action 'self';`
   - `Strict-Transport-Security`: `max-age=31536000; includeSubDomains; preload`
   - `X-Content-Type-Options`: `nosniff`
   - `Referrer-Policy`: `strict-origin-when-cross-origin`
   - `Permissions-Policy`: `camera=(), microphone=(), geolocation=()`
3. **Verify `manifest.webmanifest` & Service Worker**:
   - Confirm `https://<domain>/manifest.webmanifest` returns HTTP `200` with `application/manifest+json`.
   - Run Lighthouse (PWA + Performance + Accessibility + Best Practices + SEO) against the live HTTPS URL.

---

## 2. Packaging Options

### Option A: Bubblewrap CLI (Recommended)

```bash
npm install -g @bubblewrap/cli
mkdir esz-office-twa && cd esz-office-twa
bubblewrap init --manifest="https://app.eszoffice.com/manifest.webmanifest"
```

When prompted by `bubblewrap init`:
- **Package ID**: `com.eszoffice.app` (or your reverse-DNS domain ID)
- **App name**: `eSZ OFFICE`
- **Launcher name**: `eSZ Office`
- **Display mode**: `standalone`
- **Signing key**: Create a new Android upload keystore (`upload-keystore.jks`) or point to your existing keystore. **Back up the `.jks` file and passwords securely.**
- **Target SDK**: Ensure `targetSdkVersion` in `twa-manifest.json` / `app/build.gradle` meets Google Play's current requirement (**API 35 / Android 15** or latest required Play target SDK).

Build the signed `.aab`:

```bash
bubblewrap build
```

### Option B: PWABuilder

1. Visit `https://www.pwabuilder.com` and enter your deployed HTTPS URL.
2. Select **Package for Stores → Android**.
3. Set Package ID (`com.eszoffice.app`), enable **Google Play Signing**, and download the generated `.aab` and `assetlinks.json` snippet.

---

## 3. Digital Asset Links (`/.well-known/assetlinks.json`)

To remove the browser URL bar inside the Android TWA, Google Play must verify ownership between the Android package and your web origin:

1. Extract the SHA-256 fingerprint from your upload keystore:
   ```bash
   keytool -list -v -keystore upload-keystore.jks -alias android
   ```
2. Also copy the **App signing key certificate SHA-256 fingerprint** from **Google Play Console → Setup → App signing** once uploaded.
3. Update `/public/.well-known/assetlinks.json` (remove `_instructions` and replace `package_name` and `sha256_cert_fingerprints` with both fingerprints if using Play App Signing):
   ```json
   [
     {
       "relation": ["delegate_permission/common.handle_all_urls"],
       "target": {
         "namespace": "android_app",
         "package_name": "com.eszoffice.app",
         "sha256_cert_fingerprints": [
           "YOUR_PLAY_APP_SIGNING_SHA256",
           "YOUR_UPLOAD_KEY_SHA256"
         ]
       }
     }
   ]
   ```
4. Verify `https://<domain>/.well-known/assetlinks.json` returns `200 OK` with `Content-Type: application/json` and zero redirects.

---

## 4. Google Play Console Data Safety Form Answers

Because eSZ OFFICE processes every file locally and has no backend, analytics, or ads:

- **Does your app collect or share any of the required user data types?** → **No**
- **Is all of the user data collected by your app encrypted in transit?** → N/A (No data collected or transmitted)
- **Do you provide a way for users to request that their data is deleted?** → N/A (All data lives strictly on the user's device and can be deleted in-app or by uninstalling)
- **Privacy Policy URL**: `https://<domain>/privacy-policy`
- **Contains Ads**: **No**
- **App Access**: All functionality is available without login or special credentials.

---

## 5. Content Rating & Store Listing Copy

- **Content Rating (IARC)**: Productivity / Utility, no user-generated public feeds, no ads, no IAP → **Everyone / PEGI 3**.
- **App Title (<= 30 chars)**: `eSZ OFFICE`
- **Short Description (<= 80 chars)**:
  `Offline PDF, Document, Spreadsheet & Slides suite. Private and free.`
- **Full Description**:
  > **eSZ OFFICE** is a fast, free, offline-first office suite that keeps every file strictly on your device.
  >
  > **Four Integrated Modules**
  > - **PDF**: View, organize, merge, split, annotate, sign, encrypt, and OCR scanned PDFs locally.
  > - **Documents**: Word-style rich text editor supporting `.docx`, `.txt`, `.md`, `.html`, `.rtf`, and PDF export.
  > - **Spreadsheets**: Multi-sheet workbooks with formulas, charts, sorting, filtering, and `.xlsx`/`.xls`/`.csv` support.
  > - **Presentations**: View and build slide decks with speaker notes, themes, and `.pptx`/PDF/PNG export.
  >
  > **Privacy First**
  > - No accounts, no sign-in, no ads, and no telemetry.
  > - All processing happens on-device; your files are never uploaded.
  > - Export and restore full `.zip` backups anytime.

---

## 6. Screenshots & Release Checklist

- [ ] **App Icon**: 512×512 PNG (`icon-512.png`) + 1024×500 Feature Graphic for Play Console.
- [ ] **Phone Screenshots (minimum 2–4, 16:9 or 9:16)**:
  - Home screen showing New actions, Recent, and Offline status
  - Local File Manager with folders & thumbnails
  - Tabbed Multi-Document Workspace
  - Settings & Storage / ZIP Backup view
- [ ] **Tablet Screenshots (7-inch & 10-inch)**: Captured at `1280x800` / `1920x1200`.
- [ ] **Verify TWA Full-Screen**: Install internal test track from Play Console on a physical Android device and confirm the Chrome URL bar is hidden (`assetlinks.json` verified).
- [ ] **Verify Offline Cold Start**: Enable Airplane Mode and launch the installed app from the Android launcher.
