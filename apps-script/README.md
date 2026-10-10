# Signup form → Google Sheet

1. Create a new Google Sheet (e.g. "Mingleby Signups").
2. In the sheet: **Extensions → Apps Script**. Replace the contents of `Code.gs` with `apps-script/Code.gs` from this repo. Save.
3. **Deploy → New deployment** → gear icon → **Web app**:
   - Execute as: **Me**
   - Who has access: **Anyone**
   - Click **Deploy**, approve the permissions prompt (Advanced → Go to project → Allow).
4. Copy the **Web app URL** (ends in `/exec`). Open it in a browser — you should see `{"ok":true,...}`.
5. Paste the URL into `index.html`, replacing `PASTE_YOUR_APPS_SCRIPT_WEB_APP_URL_HERE` (search for `SHEET_URL`).

The `Signups` tab and header row are created automatically on the first submission.

## Updating the script later

After editing `Code.gs`, use **Deploy → Manage deployments → ✏️ Edit → Version: New version → Deploy**.
This keeps the same `/exec` URL. Creating a *new deployment* instead gives a new URL that you'd need to paste again.

## How rows work

| Step | What happens |
|---|---|
| Name + email submitted | New row: `lead_id`, name, partner, email, source, UTMs, `step2_status = pending` |
| Interests submitted | Same row (matched by `lead_id`) gets `interests`, `city`, `step2_status = completed` |
| "Skip for now" | Same row gets `step2_status = skipped` |

If step 2 can't find its row, it appends a full new row with name and email included.
