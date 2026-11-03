# MEAL-Dashboard

Don Bosco Kakuma MEAL learner and employment monitoring app, built with Streamlit.

## Local run

```powershell
python -m pip install -r requirements.txt
streamlit run app.py
```

Without `DATABASE_URL`, the app uses a local `meal.db` SQLite database.

Create an account once, then use **Login** on later visits. The login form's
**Keep me signed in** option remembers the current device until you log out.
Use it only on a private device; browsers may clear or limit stored cookies.
Signing out revokes that device's remembered session.

New student admission numbers are generated as `CENTRE/COURSE###/YEAR`
(for example, `DBK/COMP001/2026`). Numbering starts at 001 for each centre,
course and year. Student-register uploads keep supplied admission numbers and
generate numbers for new rows that do not have one. Administrators can generate
numbers for existing records with missing identifiers from the learner records
page; existing admission numbers are not changed.

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

The statistical report generator uses the data-free `END SEPTEMBER 2026 STATISTICAL REPORT.xlsx` template included in the repository. Other Excel workbooks remain excluded from Git. Never commit or upload a workbook containing identifiable student records.

### Database behavior

- Local mode (no `DATABASE_URL`): SQLite in `meal.db`.
- Hosted mode (`DATABASE_URL` set): PostgreSQL; tables and indexes are initialized at app startup.
- The two databases are independent. No existing local learner or employment data is migrated automatically.
- Remembered sign-in sessions, audit logs, and admission-number sequences remain in the configured database. Keep `DATABASE_URL` pointed at the same persistent database between deployments.
- When Google Sheets is configured, account profiles and PBKDF2-SHA256 password hashes are stored in the `User Accounts` tab. Without Google Sheets credentials, account profiles and hashes remain in the configured database.
- Learner and employment records can use Google Sheets when the `gcp_service_account` secret is configured.
- Keep `meal.db`, report workbooks containing personal data, and `.streamlit/secrets.toml` out of Git.
- New account passwords are stored using PBKDF2-SHA256. Existing legacy password hashes are upgraded after a successful login.

## Google Sheets records and accounts

The app uses the `Learners` and `Employment` tabs for learner and graduate-tracer records. When Sheets is configured, it creates a `User Accounts` tab and uses it for sign-ups, logins, account approvals, and password hashes. Remembered sign-in sessions, audit logs, and admission-number sequences remain in the configured database.

1. Enable the Google Sheets API in the Google Cloud project associated with the service account.
2. Add the service-account JSON fields to Streamlit Community Cloud **App settings → Secrets**. Keep the existing `DATABASE_URL` secret, then add:

   ```toml
   [google_sheets]
   spreadsheet_id = "<spreadsheet-id>"

   [gcp_service_account]
   # Copy the service-account JSON fields here locally; never commit credentials.
   ```

   Include the remaining fields from the service-account JSON file as well. Never commit the credentials or paste them into chat or source code.
3. Share the spreadsheet with the `client_email` value as an **Editor**. Keep the sheet private and grant access only to authorized staff and the service account.
4. Deploy/restart the app. On first connection, it verifies the `Learners` and `Employment` tabs, appends missing columns without removing existing data, creates `User Accounts`, and copies existing database account rows there once. The original account rows are retained in the database as a rollback backup. New sign-ups, approvals, and password-hash upgrades then use the sheet. Existing learner and employment database records are not migrated automatically.

**Security:** The `User Accounts` tab contains names, email addresses, roles, approval status, and PBKDF2 password hashes. Passwords are never stored as plain text, but spreadsheet editors can see the hashes and can change roles or approvals. Restrict sharing to trusted administrators and the app’s service account; anyone with sheet access may view or alter all its data. Remember-me session tokens remain in the configured database.

### Account approval emails

The app sends an approval confirmation to the address used during sign-up after an administrator approves the account. The first administrator account is approved immediately and receives the same confirmation. If email delivery fails, the account approval remains successful and the app displays a warning.

The app sends email through Gmail SMTP. Turn on 2-Step Verification for the sender's Google account, then create a Google App Password. Add the Gmail address and app password under Streamlit Community Cloud **App settings → Secrets**:

```toml
[email]
gmail_address = "your-account@gmail.com"
gmail_app_password = "your 16-character app password"
```

Keep the app password private. Do not commit it to GitHub or paste it into chat. Local development may use the same `[email]` section in the ignored `.streamlit/secrets.toml`. Admins can use **User Approvals → Test or resend account email** to test delivery to their own account or resend an approval notice. The hosting environment must allow outbound SMTP over port 587; if it blocks SMTP, Gmail delivery cannot work from that host. Check inbox and spam folders for test messages.

### Admin notifications

Administrators can use **Notifications** in the sidebar to review account sign-ups, approvals and rejections, learner and employment record changes, imports, and bulk deletions. Notifications are stored in the database until an administrator marks them read. Read status is shared across administrator accounts. The sidebar shows the total unread count; the panel provides per-notification and mark-all actions, with older notifications available through pagination.

For local development, add the same `[google_sheets]` and `[gcp_service_account]` sections to the ignored `.streamlit/secrets.toml`. Do not put credentials in the repository. When credentials are absent, accounts and learner/employment records remain in the configured database. Use test records first and confirm the spreadsheet is shared with the service account before entering personal data.

### Nationality values

Learner registration and employment outcome forms use the nationality labels provided by the program, including `Kenyan (Others)` and `Kenyan (Host Community)`. Nationality is saved with both record types. Existing Kenyan and Turkana values remain recognized for reports; Turkana is counted in the Kenyan Host Community category.

### Duplicate learner names

The app treats learner names as duplicates when they match after trimming and collapsing whitespace and ignoring letter case. On an administrator session, it removes existing duplicate learner records from the active learner store once, keeps the earliest-created record, and records the deletions in Change History. Registration and bulk upload reject names that already exist, including repeated names within one upload. Since different learners can share a name, review names carefully before enabling this cleanup on existing records.