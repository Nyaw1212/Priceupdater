/**
 * Master Material Sheet price cross-check.
 *
 * Match: Sheet1 column D <-> Price updater column A
 * Write: Price updater column B -> Sheet1 column Z
 *        Price updater column A -> Sheet1 column AA
 */
const CONFIG = {
  masterSheet: 'Sheet1',
  priceUpdaterSheet: 'Price updater',
  startRow: 2,

  masterCatalogColumn: 4, // D
  outputStartColumn: 26, // Z
  updaterCatalogColumn: 1, // A
  updaterValueColumn: 2 // B
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Price Update')
    .addItem('Run cross-check', 'runPriceCrossCheck')
    .addToUi();
}

function runPriceCrossCheck() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = spreadsheet.getSheetByName(CONFIG.masterSheet);
  const priceUpdaterSheet = spreadsheet.getSheetByName(CONFIG.priceUpdaterSheet);

  if (!masterSheet || !priceUpdaterSheet) {
    throw new Error(
      'Expected sheets "' + CONFIG.masterSheet + '" and "' +
        CONFIG.priceUpdaterSheet + '". Please confirm their names.'
    );
  }

  const masterLastRow = getLastNonEmptyRow_(
    masterSheet,
    CONFIG.masterCatalogColumn,
    CONFIG.startRow
  );
  const updaterLastRow = getLastNonEmptyRow_(
    priceUpdaterSheet,
    CONFIG.updaterCatalogColumn,
    CONFIG.startRow
  );

  if (masterLastRow < CONFIG.startRow) {
    throw new Error('No catalog numbers found in Sheet1 column D.');
  }
  if (updaterLastRow < CONFIG.startRow) {
    throw new Error('No catalog numbers found in Price updater column A.');
  }

  const masterRowCount = masterLastRow - CONFIG.startRow + 1;
  const updaterRowCount = updaterLastRow - CONFIG.startRow + 1;

  const masterCatalogs = masterSheet
    .getRange(CONFIG.startRow, CONFIG.masterCatalogColumn, masterRowCount, 1)
    .getDisplayValues();

  const updaterRange = priceUpdaterSheet.getRange(
    CONFIG.startRow,
    CONFIG.updaterCatalogColumn,
    updaterRowCount,
    2
  );
  const updaterDisplay = updaterRange.getDisplayValues();
  const updaterValues = updaterRange.getValues();

  // Key: normalised catalog number. Value: original A/B values.
  // When Price updater has duplicates, the first row is used.
  const updaterByCatalog = new Map();
  const updaterCatalogs = new Set();
  const duplicatesInUpdater = [];

  updaterDisplay.forEach((row, index) => {
    const key = normaliseCatalog_(row[0]);
    if (!key) return;

    updaterCatalogs.add(key);

    if (updaterByCatalog.has(key)) {
      duplicatesInUpdater.push(row[0]);
      return;
    }

    updaterByCatalog.set(key, {
      catalog: updaterValues[index][0],
      price: updaterValues[index][1]
    });
  });

  const output = [];
  const masterCatalogKeys = new Set();
  const missingFromPriceUpdater = [];
  let catalogRowsChecked = 0;
  let matchedRows = 0;

  masterCatalogs.forEach((row) => {
    const shownCatalog = row[0];
    const key = normaliseCatalog_(shownCatalog);

    if (!key) {
      output.push(['', '']);
      return;
    }

    catalogRowsChecked++;
    masterCatalogKeys.add(key);

    const match = updaterByCatalog.get(key);

    if (match) {
      // Z receives Price updater column B; AA receives Price updater column A.
      output.push([match.price, match.catalog]);
      matchedRows++;
    } else {
      output.push(['', '']);
      missingFromPriceUpdater.push(shownCatalog);
    }
  });

  masterSheet
    .getRange(CONFIG.startRow, CONFIG.outputStartColumn, masterRowCount, 2)
    .setValues(output);

  const missingFromMaster = updaterDisplay
    .map((row) => row[0])
    .filter((catalog) => {
      const key = normaliseCatalog_(catalog);
      return key && !masterCatalogKeys.has(key);
    });

  SpreadsheetApp.flush();

  showReport_({
    catalogRowsChecked: catalogRowsChecked,
    matchedRows: matchedRows,
    missingFromPriceUpdater: uniqueSorted_(missingFromPriceUpdater),
    missingFromMaster: uniqueSorted_(missingFromMaster),
    duplicatesInUpdater: uniqueSorted_(duplicatesInUpdater)
  });

  spreadsheet.toast(
    matchedRows + ' matched; ' + missingFromPriceUpdater.length + ' not matched.',
    'Price Update Complete',
    8
  );
}

function normaliseCatalog_(catalog) {
  return String(catalog || '')
    .trim()
    .toUpperCase()
    .replace(/\s+/g, '');
}

function getLastNonEmptyRow_(sheet, column, startRow) {
  const sheetLastRow = sheet.getLastRow();
  if (sheetLastRow < startRow) return 0;

  const values = sheet
    .getRange(startRow, column, sheetLastRow - startRow + 1, 1)
    .getDisplayValues();

  for (let index = values.length - 1; index >= 0; index--) {
    if (String(values[index][0]).trim()) {
      return startRow + index;
    }
  }

  return 0;
}

function uniqueSorted_(values) {
  return [...new Set(values)].sort();
}

function showReport_(report) {
  const html = `
    <div style="font-family:Arial,sans-serif;padding:16px;color:#202124;">
      <h2 style="margin:0 0 16px;">Price Update Cross-check</h2>
      <p><strong>Catalog rows checked:</strong> ${report.catalogRowsChecked}</p>
      <p><strong>Matched:</strong> ${report.matchedRows}</p>
      <p><strong>Not found in Price updater:</strong> ${report.missingFromPriceUpdater.length}</p>
      <p><strong>Price updater catalog numbers not found in Sheet1:</strong> ${report.missingFromMaster.length}</p>
      <p><strong>Duplicate catalog numbers in Price updater:</strong> ${report.duplicatesInUpdater.length}</p>
      ${makeListHtml_('Not found in Price updater', report.missingFromPriceUpdater)}
      ${makeListHtml_('Price updater numbers not found in Sheet1', report.missingFromMaster)}
      ${makeListHtml_('Duplicates in Price updater (first occurrence used)', report.duplicatesInUpdater)}
    </div>
  `;

  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(680).setHeight(620),
    'Price Update Report'
  );
}

function makeListHtml_(title, catalogNumbers) {
  const list = catalogNumbers.length
    ? catalogNumbers.map(escapeHtml_).join('\n')
    : 'None';

  return `
    <h3 style="margin:20px 0 8px;">${escapeHtml_(title)}</h3>
    <pre style="max-height:135px;overflow:auto;margin:0;padding:8px;white-space:pre-wrap;background:#f8f9fa;border:1px solid #dadce0;">${list}</pre>
  `;
}

function escapeHtml_(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}
