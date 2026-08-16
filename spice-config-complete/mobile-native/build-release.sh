#!/usr/bin/env bash
# Build a distributable (signed if configured) release APK for the wrapper.
# Run from the mobile-native/ directory. See README for one-time signing setup.
set -euo pipefail
cd "$(dirname "$0")"

if [ ! -d android ]; then
  echo "No android/ project yet — creating it..."
  npm install
  npx cap add android
fi

npm install
npx cap sync android

echo "Building release APK..."
( cd android && ./gradlew assembleRelease )

APK="android/app/build/outputs/apk/release/app-release.apk"
[ -f "$APK" ] || APK="android/app/build/outputs/apk/release/app-release-unsigned.apk"
echo ""
echo "Done. APK at: mobile-native/$APK"
echo "Copy it to the phone and install (allow install from unknown sources)."
