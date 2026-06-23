# Exhibition Match Maker Pilipinas

Static web UI for a billiard matchmaking platform using an Excel workbook as the database.

## Files

- `index.html` - main web app
- `assets/styles.css` - styling
- `assets/app.js` - Excel import/export and UI logic
- `data/BilliardMatchDatabase.xlsx` - Excel database

## Free deployment with free domain

1. Create a free GitHub account.
2. Create a public repository named `exhibition-match-maker-pilipinas`.
3. Upload all files in this folder to the repository root.
4. Go to **Settings > Pages**.
5. Under **Build and deployment**, choose **Deploy from a branch**.
6. Select the `main` branch and `/root`, then Save.
7. Your free domain will be:
   `https://YOUR-GITHUB-USERNAME.github.io/exhibition-match-maker-pilipinas/`

## Excel-only database workflow

This is a static site. Browser users can add profiles and match requests, then export `BilliardMatchDatabase.xlsx`.

To publish new data for everyone:

1. Open the site.
2. Add/import/update data.
3. Click **Export Updated Excel**.
4. Replace `data/BilliardMatchDatabase.xlsx` in GitHub with the exported file.
5. GitHub Pages will serve the updated workbook.

## Important limitation

Real Facebook OAuth login, SMS verification, private accounts, and server-side signup require a backend. This MVP keeps maintenance low by using a static site and Excel-only data. It stores temporary changes in browser localStorage and exports Excel.
