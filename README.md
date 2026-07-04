# SecureLabX Company Portal

Static site with company info, a Leave Request form, and a Claim Request form.
Protected by a shared password gate. Form submissions are stored in a Google
Sheet via a small Google Apps Script backend.

## Security note

The password gate is **client-side only**. It stops casual visitors, but
anyone who inspects the page source or brute-forces the password hash can get
past it — there is no real per-user authentication. Don't put confidential
data behind this alone. If you need stronger access control later (e.g. a
real per-email allowlist), consider adding Cloudflare Access in front of the
GitHub Pages site.

## 1. Set your own password

1. Open `index.html` (or any page) in a browser, or just open DevTools
   console anywhere.
2. Run this in the console, replacing `yourNewPassword`:
   ```js
   crypto.subtle.digest("SHA-256", new TextEncoder().encode("yourNewPassword"))
     .then(buf => console.log([...new Uint8Array(buf)].map(b => b.toString(16).padStart(2, "0")).join("")));
   ```
3. Copy the printed hash into `assets/js/config.js` as `SITE_PASSWORD_HASH`.

The default password (until you change it) is `SecureLabX@2026`.

## 2. Set up the Google Sheet + Apps Script backend

1. Go to [sheets.google.com](https://sheets.google.com) and create a new
   spreadsheet, e.g. "SecureLabX Submissions".
2. In the sheet, open **Extensions > Apps Script**.
3. Delete any starter code and paste in the contents of
   [`apps-script/Code.gs`](apps-script/Code.gs) from this repo.
4. Click **Deploy > New deployment**.
   - Select type: **Web app**.
   - Execute as: **Me**.
   - Who has access: **Anyone**.
5. Click **Deploy**, authorize the permissions Google asks for, then copy the
   **Web app URL** it gives you.
6. Paste that URL into `assets/js/config.js` as `APPS_SCRIPT_URL`.

Submissions will automatically create a "Leave Request" sheet tab and a
"Claim Request" sheet tab the first time each form is used, with a timestamp
column plus one column per form field.

If you edit deployed script code later, use **Deploy > Manage deployments >
Edit > New version** so the live URL picks up your changes.

## 3. Host on GitHub Pages

From this folder:

```bash
git init
git add .
git commit -m "Initial SecureLabX portal"
git branch -M main
git remote add origin https://github.com/<your-username>/<your-repo>.git
git push -u origin main
```

Then on GitHub:

1. Go to the repo's **Settings > Pages**.
2. Under "Build and deployment", set **Source** to `Deploy from a branch`.
3. Choose branch `main`, folder `/ (root)`, then **Save**.
4. Your site will be live at `https://<your-username>.github.io/<your-repo>/`
   within a minute or two.

## File structure

```
index.html              Company info / home page
leave-request.html       Leave Request form
claim_form/claim_form.html  Claim Request form (staff claim system)
assets/css/style.css     Shared styling
assets/js/config.js      Password hash + Apps Script URL (edit these)
assets/js/auth.js         Password gate logic
assets/js/form-submit.js  Form submission handling
apps-script/Code.gs      Backend script for Google Sheets storage
```

## Adding another form

1. Copy `leave-request.html` to a new file.
2. Update the nav links on every page to include the new page.
3. Change the hidden `formType` input value and the form fields.
4. The Apps Script backend will auto-create a new sheet tab named after the
   `formType` value the first time it's submitted.
