# Tharaa Audit Fixing Dashboard

Optimizers-branded static dashboard for reviewing, filtering, and tracking Tharaa audit fixes.

## What Is Included

- `index.html` is the deployable dashboard.
- `scripts/build_tharaa_audit_dashboard.py` is the local rebuild script used to regenerate the dashboard from the workbook.
- `vercel.json` tells Vercel to serve the static dashboard.

## Deploy To Vercel

1. Push this folder to a GitHub repository.
2. In Vercel, create a new project from that repository.
3. Keep the framework preset as `Other`.
4. Keep the build command as `npm run build`.
5. Keep the output directory as `.`.
6. Deploy.

## Push To GitHub

```bash
git add .
git commit -m "Initial Tharaa audit fixing dashboard"
git branch -M main
git remote add origin https://github.com/YOUR-USERNAME/YOUR-REPO.git
git push -u origin main
```

## Local Preview

Open `index.html` directly in a browser, or run:

```bash
npm run dev
```

Then open `http://localhost:3000`.

The dashboard is static and stores fixed/not-fixed status and notes in the browser's local storage.

## Updating The Dashboard

If the workbook changes, run:

```bash
python scripts/build_tharaa_audit_dashboard.py
```

To use a different workbook path:

```bash
THARAA_AUDIT_WORKBOOK="/path/to/workbook.xlsx" python scripts/build_tharaa_audit_dashboard.py
```

The script regenerates `index.html`.
