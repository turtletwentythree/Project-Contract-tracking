# T23 Tracking Contract

Production static dashboard for T23 Contract Tracking.

## Production

This repository is intended to run on GitHub Pages from the root `index.html`.

## Files

- `index.html` - production dashboard
- `data/contracts_db.csv` - contract database export
- `data/log_db.csv` - log database export
- `data/type_master_db.csv` - contract type master data
- `src/tracking_contracts_dashboard_codex_code.py` - Codex generator source
- `src/attachment_upload_apps_script.js` - Google Apps Script Web App source for email and attachments

## Notes

The dashboard uses browser local storage for local contract data changes and the deployed Google Apps Script endpoint for email/attachment handling.

## Production Operations

After changing `src/attachment_upload_apps_script.js`, update the deployed Google Apps Script Web App with the same source and redeploy it.

To enable daily backup after redeploying Apps Script, send a POST request to the Web App with:

```json
{
  "mode": "installDailyBackup"
}
```

The backup trigger runs `runDailyBackup()` once per day and writes CSV copies with the `daily_backup_` prefix to the configured backup folder.
