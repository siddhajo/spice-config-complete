# Testing the native wrapper (Mac & Windows)

Bluetooth printing must be tested on a **real Android phone** with the printer
**paired** — an emulator can't talk to a physical printer. The build steps are
the same on Mac and Windows except where noted.

---

## 0. One-time installs (both platforms)

1. **Node.js LTS (18+)** — <https://nodejs.org> (download the LTS installer).
   - Mac (alternative): `brew install node`
   - Verify: `node -v` and `npm -v`
2. **Android Studio** — <https://developer.android.com/studio>. Install, open it
   once, and let it finish downloading the Android SDK.
   - It **includes the correct JDK (17)** — you do NOT need a separate Java
     install if you build from Android Studio (recommended below).
3. A **USB cable** for the phone.

### Phone prep (do once)
- **Enable Developer mode:** Settings → About phone → tap **Build number** 7
  times. Then Settings → System → Developer options → turn on **USB debugging**.
- **Pair the printer:** Settings → Bluetooth → pair the thermal printer (Classic
  printers must be paired at the OS level before the app can use them).

---

## 1. Configure (both platforms)

Open a terminal in the `mobile-native` folder:
- **Mac:** Terminal.app
- **Windows:** PowerShell **or** Git Bash

```bash
npm install
```

Then edit **`capacitor.config.json`** and set `server.url` to your live host,
e.g. `https://your-app.up.railway.app/mobile` — replacing the placeholder with
the real one. (Make sure the updated `public-mobile/app.html` is deployed to
that host.)

Create the Android project and sync the plugin:

```bash
npx cap add android
npx cap sync
```

---

## 2. Build + install on the phone — easiest path (Mac & Windows identical)

```bash
npx cap open android
```

This opens the project in **Android Studio**. Then:
1. Plug the phone in via USB (accept the "Allow USB debugging?" prompt on the
   phone).
2. In Android Studio's toolbar, the phone should appear as the target device
   (top center). If not, click the device dropdown and select it.
3. Click the green **Run ▶** button.
4. Android Studio builds, installs, and launches the app on the phone. Accept
   the **Nearby devices / Bluetooth** permission prompt the first time.

> This GUI path uses Android Studio's bundled JDK, so you avoid all
> `JAVA_HOME`/Gradle path issues. Use it for testing.

---

## 3. Build an APK from the command line (optional)

Use this to produce a shareable `.apk` file instead of Run-on-device.

**Mac (or Windows Git Bash):**
```bash
./build-release.sh
# or a quick debug build:
cd android && ./gradlew assembleDebug
```

**Windows (PowerShell / Command Prompt):**
`build-release.sh` is a bash script — in PowerShell/CMD run the steps directly:
```powershell
npx cap sync android
cd android
.\gradlew.bat assembleDebug
```

Resulting file:
`mobile-native/android/app/build/outputs/apk/debug/app-debug.apk`
(release build → `.../apk/release/app-release.apk`).

Copy that APK to the phone (USB, email, or Drive) and tap it to install — allow
**"install from unknown sources"** if prompted.

---

## 4. Verify it works (on the phone)

1. Open the **Spice Auction** app and log in.
2. Menu (⋮) → **Print method** → cycle to **Bluetooth (app)**.
3. Menu → **Connect Bluetooth printer** → pick the printer.
4. Menu → **🧪 Test print** → a small **"PRINTER OK"** slip should print.
5. Now print a real lot / seller / batch — it goes straight to the printer, **no
   RawBT, no watermark**.

Worth covering all four print paths once, since each hits a different endpoint:

| Action | Endpoint exercised |
|---|---|
| Print one lot (after save) | `/api/lots/:id/receipt.escpos` |
| 📋 All by Seller | `/api/lots/print-seller.escpos` |
| My Lots → 🖨 Print selected | `/api/lots/print-batch.escpos` |
| All sellers in the trade | `/api/lots/print-all-sellers-escpos/:id` |

The last one is the important stress test: it's the biggest job, and it's where
a printer that drops bytes on large writes would show up as a truncated slip.

---

## Troubleshooting

| Symptom | Fix |
|---|---|
| Phone not shown in Android Studio | USB debugging on? Try another cable/port; tap "Allow" on the phone. |
| `npx cap` not found | Run `npm install` again inside `mobile-native`. |
| Gradle/Java error on CLI build | Build from Android Studio instead (uses its bundled JDK). |
| "Print method: Bluetooth (app)" option missing | The wrapper's bridge didn't load — rebuild after `npx cap sync`; if a remote `server.url` doesn't surface it, use the **local-bundle fallback** in README.md. |
| "No paired printers" | Pair the printer in Android Bluetooth settings first, and grant "Nearby devices". |
| Columns wrap / text runs off the edge | The roll isn't 32 characters wide. Try `?width=42` or `?width=48` on the escpos endpoints. |
| App shows a blank/placeholder page | `server.url` is wrong/unreachable, or the phone has no internet. |
