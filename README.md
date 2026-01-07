
# Attachment Size Smart Alert (Outlook add‑in)

Warn users **before send** when the **total size of attachments** exceeds a configurable threshold **x**.  
Works with **Outlook on the web**, **new Outlook for Windows**, and **classic Outlook for Windows** (event‑based “Smart Alerts”).

> **What it does**
> - Sums file and item attachments in a composed email
> - Applies a configurable **transport overhead** (+33%) to reflect MIME/Base64 expansion
> - Shows a dialog with **Cancel** (block send) or **Send anyway** (allow)
> - Optionally **excludes cloud attachments** (OneDrive/SharePoint links) from the total

---

## 📂 Project structure
outlook-smart-alerts-attachments/
├─ manifest/
│  └─ manifest.xml
├─ src/
│  ├─ runtime/
│  │  ├─ commands.html
│  │  └─ commands.js
│  └─ dialog/
│     ├─ dialog.html
│     └─ dialog.js
└─ assets/
├─ icon-32.png
├─ icon-64.png
└─ icon-80.png

- **`manifest.xml`** – add‑in manifest (XML) with event registration (`OnMessageSend`) and HTTPS resource URLs.  
- **`commands.html` / `commands.js`** – event runtime; calculates size and launches the dialog.  
- **`dialog.html` / `dialog.js`** – user prompt to cancel or proceed.  
- **`assets/`** – icons shown in add‑in listings.

---

## 🚀 Hosting (GitHub Pages)

This repo is designed to be hosted via **GitHub Pages**:

1. Enable Pages: **Settings → Pages → Build and deployment**  
   - Source: `Deploy from a branch`  
   - Branch: `main`  
   - Folder: `/ (root)`
2. Your site will be available at:  
   `https://ccdanc.github.io/outlook-smart-alerts-attachments/`
3. Ensure these URLs resolve:
   - `/src/runtime/commands.html`
   - `/src/dialog/dialog.html`
   - `/assets/icon-64.png`

The manifest already points to the above GitHub Pages URLs.

---

## ⚙️ Configuration

Open `src/runtime/commands.js` and adjust:

```js
const THRESHOLD_MB = 20;      // Your guidance threshold (x)
const OVERHEAD_FACTOR = 1.33; // MIME/Base64 overhead (~33%)
const EXCLUDE_CLOUD_ATTACHMENTS = true; // true = ignore OneDrive/SharePoint links
