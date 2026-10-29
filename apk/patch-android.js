// Dijalankan setelah `npx cap add android`: izin penyimpanan, tanda tangan tetap, nomor versi.
const fs = require("fs");
const path = require("path");

const android = path.join(__dirname, "android");
const manifestPath = path.join(android, "app/src/main/AndroidManifest.xml");
const gradlePath = path.join(android, "app/build.gradle");
const run = process.env.RUN_NUMBER || "1";

let m = fs.readFileSync(manifestPath, "utf8");
if (!m.includes("READ_EXTERNAL_STORAGE")) {
  m = m.replace(
    "<application",
    '<uses-permission android:name="android.permission.READ_EXTERNAL_STORAGE" />\n' +
      '    <uses-permission android:name="android.permission.WRITE_EXTERNAL_STORAGE" android:maxSdkVersion="32" />\n    <application'
  );
}
if (!m.includes("requestLegacyExternalStorage")) {
  m = m.replace("<application", '<application android:requestLegacyExternalStorage="true"');
}
fs.writeFileSync(manifestPath, m);

let g = fs.readFileSync(gradlePath, "utf8");
g = g.replace(/versionCode\s+\d+/, "versionCode " + run).replace(/versionName\s+"[^"]*"/, 'versionName "1.0.' + run + '"');
if (!g.includes("signingConfigs.rekapyuk")) {
  g += `
android {
    signingConfigs {
        rekapyuk {
            storeFile file("../../rekapyuk.keystore")
            storePassword "rekapyuk-sideload"
            keyAlias "rekapyuk"
            keyPassword "rekapyuk-sideload"
        }
    }
    buildTypes {
        release { signingConfig signingConfigs.rekapyuk }
        debug { signingConfig signingConfigs.rekapyuk }
    }
}
`;
}
fs.writeFileSync(gradlePath, g);
console.log("Patch Android selesai, versionCode", run);
