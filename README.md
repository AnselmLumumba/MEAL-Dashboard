# MEAL-Dashboard

Don Bosco Kakuma MEAL learner and employment monitoring app, built with Streamlit.

## Local run

```powershell
python -m pip install -r requirements.txt
streamlit run app.py
```

Without `DATABASE_URL`, the app uses a local `meal.db` SQLite database.

## Streamlit Community Cloud + Supabase

The app uses PostgreSQL when `DATABASE_URL` is configured; otherwise, it uses local SQLite. The hosted database is initialized with an empty schema on first startup. The local `meal.db` is not uploaded or copied to the cloud.

1. Create a Supabase project and a strong database password. In the Supabase dashboard, open **Connect** and copy a PostgreSQL connection string. For Streamlit Community Cloud, use the Supabase shared pooler connection if the direct endpoint is not reachable over IPv6. Keep SSL enabled.
2. Deploy the GitHub repository to Streamlit Community Cloud using `app.py`. Set the app to **Private** and invite only authorized staff before making it available; the first registered account becomes administrator. Then add this secret in the app's settings:

   ```toml
   DATABASE_URL = "postgresql://<database-user>:<database-password>@<database-host>:<port>/<database-name>?sslmode=require"
   ```

   Use the exact connection string supplied by Supabase, replacing its password placeholder. Never commit the real URL or paste it into source code, issues, or chat.
3. Open the private app and create the first account. Approve centre-head accounts only for authorized staff. Use a strong Supabase database password and keep the database credentials private.
4. Use test records first. The hosted Supabase database is separate from the local SQLite database and starts empty. Do not upload real student data until staff access, backups, data-protection requirements, and database recovery have been reviewed.

### Excel report template

The statistical report generator looks for `END SEPTEMBER 2026 STATISTICAL REPORT.xlsx`. Excel workbooks are excluded from Git. To enable hosted report generation, make a reviewed, data-free template available in the deployment and configure `REPORT_TEMPLATE_PATH` to its path. Without that template, the dashboard still runs but Excel report generation is unavailable. Do not deploy a workbook containing identifiable student records.

### Database behavior

- Local mode (no `DATABASE_URL`): SQLite in `meal.db`.
- Hosted mode (`DATABASE_URL` set): PostgreSQL; tables and indexes are initialized at app startup.
- The two databases are independent. No existing local learner or employment data is migrated automatically.
- Keep `meal.db`, report workbooks containing personal data, and `.streamlit/secrets.toml` out of Git.
- New account passwords are stored using PBKDF2-SHA256. Existing legacy password hashes are upgraded after a successful login.