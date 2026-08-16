# Spice Auction — native Android wrapper

A **thin** native Android app (Capacitor) that loads the existing mobile PWA in
a WebView and adds one thing the browser can't do: **direct Bluetooth Classic
(SPP) ESC/POS printing**. That lets receipts print straight to a basic thermal
printer like the **HPRT MPT-II** — **without RawBT, and therefore without
RawBT's free-version watermark.**

Nothing about the web app changes: the WebView loads your live
`https://<host>/mobile`, so app updates are automatic (no APK rebuild needed for
normal app changes). The wrapper only supplies the Bluetooth bridge
(`cordova-plugin-bluetooth-serial`, exposed to the page as
`window.bluetoothSerial`). The web app already detects that bridge and adds a
**"Bluetooth (app)"** print method (see `public-mobile/app.html`).

> This scaffold is ready to build, but it must be **built on a machine with
> Android tooling** (or in CI — see `CLOUD-BUILD.md`, which is the easier path).

## Before anything else: set your host

Edit **`capacitor.config.json`** → `server.url` and replace the placeholder
`https://your-app.up.railway.app/mobile` with your real deployed host. The APK
is useless until this points at a reachable `/mobile`.

## Prerequisites (on the build machine)

- **Node.js** 18+
- **Android Studio** (latest) with an Android SDK installed
- **JDK 17** (bundled with recent Android Studio)

## Build steps

```bash
cd mobile-native

# 1. Install Capacitor + the Bluetooth-serial plugin
npm install

# 2. Point the app at YOUR hosted mobile URL (see above)
#    capacitor.config.json → server.url

# 3. Create the native Android project and pull in the plugin
npx cap add android
npx cap sync

# 4. Add Bluetooth permissions (see next section), then open in Android Studio
npx cap open android
```

In Android Studio: **Build → Build Bundle(s)/APK(s) → Build APK(s)**, then copy
the generated `app-debug.apk` to the phone and install it (allow "install from
unknown sources"). For distribution, create a signed release build instead.

## Android permissions (required)

Open `android/app/src/main/AndroidManifest.xml` and make sure these are present
inside `<manifest>` (the plugin adds the legacy ones; add the Android-12+ ones
if missing — the CI workflow does this automatically):

```xml
<uses-permission android:name="android.permission.BLUETOOTH" android:maxSdkVersion="30" />
<uses-permission android:name="android.permission.BLUETOOTH_ADMIN" android:maxSdkVersion="30" />
<!-- Android 12 (API 31)+ -->
<uses-permission android:name="android.permission.BLUETOOTH_CONNECT" />
<uses-permission android:name="android.permission.BLUETOOTH_SCAN" />
```

On Android 12+ the app must be granted **Nearby devices / Bluetooth** permission
at runtime — accept the prompt the first time you tap **Connect Bluetooth
printer**. If no prompt appears, enable it manually in Android **Settings → Apps
→ Spice Auction → Permissions → Nearby devices**.

## Using it (on the phone)

1. **Pair the printer first** in Android **Settings → Bluetooth** (Classic SPP
   printers must be paired at the OS level before an app can open them).
2. Open the installed **Spice Auction** app and log in.
3. Menu (⋮) → **Print method** → cycle to **Bluetooth (app)**.
4. Menu → **Connect Bluetooth printer** → pick the printer from the in-app list
   (auto-selected if it's the only paired device; remembered afterwards).
5. Menu → **🧪 Test print** → a small "PRINTER OK" slip should come out. Once
   that works, everything else does.
6. Print any lot / seller / batch — the ESC/POS goes straight to the printer.
   No RawBT, no watermark.

The web app already handles the fiddly parts so they don't come back as
problems: ESC/POS is sent in **small chunks** (big "all sellers" jobs don't get
truncated), the link is **re-established before each job** if it dropped while
idle, and the printer is picked from a **proper in-app list** (not a typed
prompt) and remembered.

## Distributable signed APK

A signed release APK installs cleanly and updates without the "unknown app"
friction of debug builds. One-time setup, then one command per build. (Or skip
all of this and use `CLOUD-BUILD.md`.)

```bash
cd mobile-native

# 1. Generate a signing key (keep the keystore + passwords safe; you need the
#    SAME key for every future update).
keytool -genkeypair -v -keystore android/spice-release.keystore \
  -alias spice -keyalg RSA -keysize 2048 -validity 10000

# 2. Tell Gradle about it
cp signing/keystore.properties.example android/keystore.properties
#    edit android/keystore.properties → set the two passwords you just chose

# 3. Wire the signing config in: add this as the LAST line of
#    android/app/build.gradle
#        apply from: "../../signing/signing.gradle"

# 4. Build
./build-release.sh
#    → mobile-native/android/app/build/outputs/apk/release/app-release.apk
```

`build-release.sh` runs `npm install`, creates the android project if needed,
`cap sync`, and `gradlew assembleRelease`. If you skip the signing setup it
still builds, but the APK is unsigned (fine for a quick test, not for handing
out).

## How it fits together

- `capacitor.config.json` → `server.url` = your live `/mobile` (WebView target).
- `cordova-plugin-bluetooth-serial` → injects `window.bluetoothSerial`
  (`list` / `connect` / `connectInsecure` / `isConnected` / `write`).
- `public-mobile/app.html` →
  - `hasNativeBt()` detects the bridge,
  - adds the **`native`** print method (and prefers it inside the wrapper),
  - `nativeBtPrint()` fetches the `….escpos` bytes and `write()`s them in chunks,
  - `connectPrinter()` (the menu item) picks and remembers the printer.
- `mobile-bridge.js` → serves the ESC/POS bytes: `/api/lots/:id/receipt.escpos`,
  `/api/lots/print-seller.escpos`, `/api/lots/print-batch.escpos`,
  `/api/lots/print-all-sellers-escpos/:auctionId`. Each is the raw-bytes twin of
  the PDF route with the same name, selecting the same lots.

## Troubleshooting

- **"Bluetooth (app)" method not offered / prints open in browser** → the
  WebView didn't get `window.bluetoothSerial`. Re-run `npx cap sync`, rebuild.
  If a remote `server.url` doesn't surface the bridge on your Capacitor version,
  use the **local-bundle fallback** below.
- **"No paired printers"** → pair the printer in Android Bluetooth settings
  first, and grant the app's "Nearby devices" permission.
- **Connects but prints garbage** → the printer's character width probably isn't
  32. The ESC/POS endpoints take `?width=42` or `?width=48`.

### Local-bundle fallback (only if the remote bridge doesn't inject)

Instead of `server.url`, bundle the web app locally and point it at the API host:

1. Copy the app in: `cp -r ../public-mobile/* www/` (replace the placeholder
   `www/index.html`).
2. In `capacitor.config.json`, **remove the `server` block**.
3. Tell the app where the API lives by injecting a global before it loads — add
   a tiny inline script in `www/app.html`'s `<head>`:
   ```html
   <script>window.SPICE_API_BASE='https://<your-real-host>';</script>
   ```
   (`app.html` already honors `window.SPICE_API_BASE`.)
4. `npx cap sync` and rebuild. Downside: you must rebuild the APK whenever the
   web app changes.
