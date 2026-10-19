# Price Updater

Google Apps Script for updating the material master price data.

## What it does

When you choose **Price Update → Run cross-check** in the Google Sheet:

- Matches `Sheet1!D:D` against `Price updater!A:A`
- Writes `Price updater!B:B` into `Sheet1!Z:Z`
- Writes the matching catalog number from `Price updater!A:A` into `Sheet1!AA:AA`
- Clears columns Z and AA when no match is found
- Displays a scrollable report of matched, unmatched, and duplicate catalog numbers

Catalog numbers are matched without regard to letter case or spaces.

## Connect this repo to the Google Sheet

```powershell
git clone https://github.com/Nyaw1212/Priceupdater.git
cd Priceupdater
npm install -g @google/clasp
clasp login
clasp create --type sheets --title "Master Material Price Updater" --parentId YOUR_SPREADSHEET_ID
clasp push
```

Then refresh the Google Sheet. On first use, Google will ask you to authorize the script.

## Update after editing

```powershell
git pull --ff-only
clasp push
```

> Do not commit `.clasp.json`. It contains the bound Apps Script project ID for your specific Google Sheet.
