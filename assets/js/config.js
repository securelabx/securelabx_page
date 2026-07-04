// SecureLabX site configuration
// 1. Change SITE_PASSWORD_HASH to your own password's SHA-256 hash.
//    Generate one by opening this site, pressing F12 (DevTools console) and running:
//      crypto.subtle.digest('SHA-256', new TextEncoder().encode('yourNewPassword'))
//        .then(buf => console.log([...new Uint8Array(buf)].map(b => b.toString(16).padStart(2, '0')).join('')))
//    Copy the printed hash below. Default password is: SecureLabX@2026
//
// 2. After deploying the Google Apps Script (see apps-script/Code.gs and README.md),
//    paste the Web App URL into APPS_SCRIPT_URL.

const SITE_CONFIG = {
  SITE_PASSWORD_HASH: "68b1c23d708db9e2fe00fcd15e73dc6f8a88b2ba10c95df9771ed93279a0f4c5",
  APPS_SCRIPT_URL: "PASTE_YOUR_GOOGLE_APPS_SCRIPT_WEB_APP_URL_HERE"
};
