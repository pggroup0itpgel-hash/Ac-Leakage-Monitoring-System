# Custom OTP Functions Setup

## 1) Install Firebase CLI

- `npm i -g firebase-tools`
- `firebase login`

## 2) Select Firebase project

- `firebase use ac-leakage-monitor-418ed`

## 3) Install dependencies

- `cd functions`
- `npm install`

## 4) Set SMTP environment variables

Use runtime env variables in your deploy environment:

- `SMTP_HOST`
- `SMTP_PORT` (587 or 465)
- `SMTP_USER`
- `SMTP_PASS`
- `SMTP_FROM` (optional)

Example (PowerShell):

- `$env:SMTP_HOST="smtp.gmail.com"`
- `$env:SMTP_PORT="587"`
- `$env:SMTP_USER="your_mail@domain.com"`
- `$env:SMTP_PASS="your_app_password"`
- `$env:SMTP_FROM="AC Leakage System <your_mail@domain.com>"`

## 5) Deploy

- From project root: `firebase deploy --only functions`

## 6) Frontend endpoint

Frontend reads this URL:

- localStorage key: `acLeakageOtpApiUrl`

For Firebase Functions:
`localStorage.setItem('acLeakageOtpApiUrl', 'https://<region>-<project>.cloudfunctions.net')`

For Google Apps Script web app:
`localStorage.setItem('acLeakageOtpApiUrl', 'https://script.google.com/macros/s/<DEPLOYMENT_ID>/exec')`