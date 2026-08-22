# study.log

A local-first React dashboard for tracking three students preparing for CUET and CLAT.

## Run it

```bash
npm install
npm run dev
```

Open the local Vite URL shown in the terminal. For a production check:

```bash
npm run build
```

## Included

- Generic daily entry form for each student with date, topics learned, topics practised, question count, remarks, study time, confidence, and next focus.
- Separate optional paper-solved section inside the same entry for exam name, paper year, and score.
- Extra progress signals: study duration, confidence level, and next focus.
- Dashboard filters for student, recent activity, team streak, practice totals, study time, and average paper score.
- Browser `localStorage` persistence, so entries remain after a refresh on the same device.
- `Connect Excel` lets the browser create and maintain a local `.xlsx` workbook. After connecting once, every saved daily entry updates the workbook automatically.
- `View spreadsheet` shows the same data as a scrollable table inside the app.
- `Connect online sheet` syncs entries across devices through the included Google Apps Script connector.

## Suggested next additions

## Connect Google Sheets

1. Create a Google Sheet and open **Extensions > Apps Script**.
2. Paste the contents of `google-apps-script.gs` into the editor and save it.
3. Deploy it as a **Web app**, execute as yourself, and allow access to anyone with the link.
4. Copy the deployment URL into the app using **Connect online sheet** on every device.

New entries will append to the `Daily log` tab, and connecting on another device imports the existing rows. The app still keeps a local browser copy as a backup. This uses Google Apps Script as the lightweight cloud connector required for cross-device storage; without some cloud connector, browsers cannot share data between machines.

To avoid entering the URL on each device, replace the empty `DEFAULT_ONLINE_SHEET_URL` value near the top of `src/App.tsx` with the `/exec` URL, then redeploy the React app. The app will connect automatically for everyone who opens the new deployment.

## Configure GitHub

The app reads the URL from `VITE_ONLINE_SHEET_URL`. For local development, copy `.env.example` to `.env.local` and add the Apps Script `/exec` URL. For GitHub Actions, add a repository **Variable** named `VITE_ONLINE_SHEET_URL`; the included build workflow passes it to Vite automatically. Do not use a secret API key here: frontend variables are included in the public browser bundle, and the Apps Script URL is an endpoint rather than a private credential.
