# Google Sheets versions (Apps Script)

## Strain collection (`strain-collection/`)
1. Make a Google Sheet and open **Extensions → Apps Script**.
2. Replace the contents of `Code.gs` with [`strain-collection/Code.gs`](strain-collection/Code.gs).
3. Add four HTML files (**+ → HTML**, name them without `.html`): `AddNewStrain`, `AddExternalStrain`,
   `OrderStrain`, `Styles`, and paste each file's contents.
4. At the top of `Code.gs`, set `backupsFolderId` (a Drive folder for backups), `labCode`,
   `alleleCode`, `labName` and, optionally, `orderSheetUrl` (the ordering sheet's address).
5. Save, reload the sheet, and run **🪱 Worm collection 🪱 → Set up blank collection sheets**.
   Approve the permissions (Advanced → Go to … (unsafe) → Allow: it's your own script).
6. Optional: run `setupDailyBackupTrigger` once from the editor for nightly backups.

Menu: create a new strain, add an existing strain (CGC lookup, prefixes for other labs), request the
selected strain, order a strain, set who gets request emails, sort newest / oldest first, backup, format.

## Ordering (`ordering/`)
1. Make a **separate** Google Sheet, open **Extensions → Apps Script**.
2. Paste [`ordering/Code.gs`](ordering/Code.gs), and add the HTML files `OrderForm` and `Styles`.
3. At the top of `Code.gs`, set `labName` and `strainSheetUrl` (your strain collection's address).
   Leave `plasmidSheetUrl` and `oligoSheetUrl` empty for now.
4. Reload and run **🧾 Lab orders 🧾 → Set up the order sheets**. Run it from the account that
   should send the status emails.
5. On the **Settings** tab, list your grants (with budgets) and the people who place orders, or use
   **Set who gets new-order emails**.

Flow: *Request an order* → orderers are emailed → mark the row **Ordered** (menu or Status dropdown)
→ the requester is emailed → **Received** → emailed again. The **Analysis** tab updates itself.
