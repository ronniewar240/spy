# Streamlit Trade Journal with Dropbox Auto-Import

This package converts your trade journal into a Streamlit app and adds a cloud-friendly Dropbox CSV import workflow.

## Run locally

```bash
pip install -r requirements.txt
streamlit run streamlit_app.py
```

## Deploy to Streamlit Community Cloud

1. Create a GitHub repository.
2. Upload these files.
3. Deploy `streamlit_app.py` in Streamlit Cloud.
4. Add secrets in Streamlit Cloud → App settings → Secrets:

```toml
[dropbox]
access_token = "YOUR_DROPBOX_ACCESS_TOKEN"
folder = "/TradeJournalExports"
```

Do **not** commit your real Dropbox token to GitHub.

## Dropbox workflow

1. Create a Dropbox folder, for example `/TradeJournalExports`.
2. Put broker CSV exports in that folder:
   - IBKR CSV
   - Wealthsimple CSV
   - NinjaTrader Performance CSV
3. In the Streamlit app, open **Dropbox auto-import**.
4. Click **Scan Dropbox Now**.
5. The app downloads new CSVs, detects the broker format, imports trades, and skips files already imported by Dropbox content hash.

## Important database note

This package still uses SQLite (`trades_streamlit.db`). SQLite is fine for local testing and demos, but Streamlit Cloud file storage can reset. For a production trading journal, migrate the database to PostgreSQL/Supabase.

## Scheduling note

This version uses a manual **Scan Dropbox Now** button because Streamlit Cloud is not a background-worker platform. For true scheduled import, use one of these:

- GitHub Actions cron calling a small Python importer
- Render/Railway scheduled worker
- Dropbox webhook receiver on a web backend

The Dropbox import logic is isolated in `dropbox_import.py`, so it can be reused later in a cron/worker.
