# Student Internal Record Portal

Beautiful local web portal for students to log in with their exam roll number, verify by OTP, and view their own Excel-backed internal assessment data.

## Excel columns

Put your workbook at:

```text
data/student_data.xlsx
```

Required columns:

- `Course Name`
- `Student's Name`
- `Roll Number`
- `Exam Roll Number`
- `Email`
- `Percentage (%)`
- `Assignment (4)`
- `Test (4)`
- `Attendance Marks (2)`
- `Total IA Marks (10)`

## Run

```powershell
python main.py
```

Open:

```text
http://127.0.0.1:8000
```

## Deploy on Render

1. Push this folder to a GitHub repository.
2. Create a new Render Web Service.
3. Use these settings:
   - Runtime: `Python 3`
   - Build Command: `pip install -r requirements.txt`
   - Start Command: `python main.py`
4. Add `RESEND_API_KEY` and `EMAIL_FROM` for real OTP emails on Render.
5. Set `ADMIN_PASSWORD` in Render for the admin panel.
6. To store uploaded Excel files in Google Drive, also set `GOOGLE_DRIVE_FOLDER_ID` and `GOOGLE_SERVICE_ACCOUNT_JSON`.

## Google Drive storage

When Google Drive is configured, Excel uploads are saved in Google Drive instead of `data/papers`, and the paper registry/selected paper are also saved in Google Drive.

1. In Google Cloud Console, enable the Google Drive API.
2. Create a service account and download its JSON key.
3. Create a Google Drive folder for portal data.
4. Share that folder with the service account email from the JSON key.
5. Set these environment variables:

```powershell
$env:GOOGLE_DRIVE_FOLDER_ID="your-drive-folder-id"
$env:GOOGLE_SERVICE_ACCOUNT_JSON='{"type":"service_account", ... }'
```

For local development, you may instead save the JSON key file outside the project and set:

```powershell
$env:GOOGLE_DRIVE_FOLDER_ID="your-drive-folder-id"
$env:GOOGLE_APPLICATION_CREDENTIALS="C:\path\to\service-account.json"
```

Do not commit the service account JSON file to GitHub.

## OTP email

Without email credentials, the app runs in demo mode and shows the generated OTP on screen.

Render free web services block outbound SMTP ports `25`, `465`, and `587`, so Gmail SMTP can work locally but fail after deployment with `Network is unreachable`. For Render, use Resend's HTTPS email API:

```powershell
$env:RESEND_API_KEY="re_xxxxxxxxx"
$env:EMAIL_FROM="Student Portal <otp@your-verified-domain.com>"
```

In the Render dashboard, set `RESEND_API_KEY` and `EMAIL_FROM` as environment variables. `EMAIL_FROM` must be a sender address verified in Resend.

For local SMTP testing, set these environment variables before running `main.py`:

```powershell
$env:SMTP_HOST="smtp.gmail.com"
$env:SMTP_PORT="587"
$env:SMTP_SECURITY="starttls"
$env:SMTP_USER="your-email@gmail.com"
$env:SMTP_PASS="your-app-password"
$env:SMTP_FROM="your-email@gmail.com"
```

For Gmail or Google Workspace, `SMTP_PASS` usually must be a Google App Password, not the normal account password. If your provider asks for SSL SMTP, use:

```powershell
$env:SMTP_PORT="465"
$env:SMTP_SECURITY="ssl"
```

When SMTP variables are set but the mail server rejects the send, the login screen now shows the mail-server error instead of silently switching to demo mode. Demo OTP is only shown when SMTP is not configured.

## Admin panel

Open:

```text
/admin
```

Default local admin password:

```text
admin123
```

For Render, set a stronger `ADMIN_PASSWORD` environment variable.

The admin panel can upload a new Excel file, save a public Google Sheet link, switch back to Excel mode, and show detected columns and row count.

For Google Sheets, share the sheet as "Anyone with the link can view". Paste the normal Google Sheet URL; the app converts it to CSV automatically.

Google Sheet papers are live links. If you edit the Google Sheet later, the student portal reads the updated Sheet data the next time a student requests/verifies OTP.

Important for Render: uploaded Excel files can disappear on redeploy unless you use Google Drive storage or add a persistent disk.
