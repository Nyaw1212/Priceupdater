/**
 * Master Material Sheet price cross-check.
 *
 * Match: Sheet1 D (catalog), F (description) vs Price updater A, B.
 * Write: Price updater C (price) -> Sheet1 Z;
 *        Price updater A (catalog) -> Sheet1 AA; tier -> AB.
 */
const CONFIG = {
  masterSheet: 'Sheet1',
  priceUpdaterSheet: 'Price updater',
  startRow: 2,

  masterCatalogColumn: 4, // D
  masterDescriptionColumn: 6, // F
  outputStartColumn: 26, // Z
  updaterCatalogColumn: 1, // A
  updaterDescriptionColumn: 2, // B
  updaterValueColumn: 3 // C
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Price Update')
    .addItem('Run cross-check', 'runPriceCrossCheck')
    .addToUi();
}

function runPriceCrossCheck() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const master = ss.getSheetByName(CONFIG.masterSheet);
  const updater = ss.getSheetByName(CONFIG.priceUpdaterSheet);
  if (!master || !updater) throw new Error('Missing Sheet1 or Price updater sheet.');
  const mLast = getLastNonEmptyRow_(master, CONFIG.masterCatalogColumn, CONFIG.startRow);
  const uLast = getLastNonEmptyRow_(updater, CONFIG.updaterCatalogColumn, CONFIG.startRow);
  if (mLast < 2 || uLast < 2) throw new Error('No catalog numbers found in the required columns.');
  const mCount = mLast - 1, uCount = uLast - 1;
  const mData = master.getRange(2, 4, mCount, 3).getDisplayValues();
  const uRange = updater.getRange(2, 1, uCount, 3);
  const uDisplay = uRange.getDisplayValues(), uRaw = uRange.getValues();
  const masters = mData.map((r, i) => ({ i, catalog: r[0], key: normaliseCatalog_(r[0]), desc: normaliseDescription_(r[2]) })).filter(x => x.key);
  const sources = uDisplay.map((r, i) => ({
    i, key: normaliseCatalog_(r[0]), desc: normaliseDescription_(r[1]), catalog: uRaw[i][0], price: uRaw[i][2]
  })).filter(x => x.key);
  const output = Array.from({ length: mCount }, () => ['', '', '']);
  const usedM = new Set(), usedU = new Set();
  const counts = { 100: 0, 95: 0, 90: 0, 80: 0, 40: 0 };
  const tiers = [100, 95, 90, 80, 40];
  const exact = new Map();
  sources.forEach(u => {
    if (!exact.has(u.key)) exact.set(u.key, []);
    exact.get(u.key).push(u);
  });
  masters.forEach(m => {
    const bucket = exact.get(m.key);
    if (!bucket || !bucket.length) return;
    const u = bucket.shift();
    usedM.add(m.i); usedU.add(u.i);
    output[m.i] = [u.price, u.catalog, '100%'];
    counts[100]++;
  });
  // Exact catalog matches take precedence, regardless of description wording.
  // Fuzzy passes combine catalog (65%) and description (35%) similarity.
  // A missing description is not allowed to create a fuzzy match.
  const remainingM = masters.filter(m => !usedM.has(m.i));
  const remainingU = sources.filter(u => !usedU.has(u.i));
  const candidates = [];
  remainingM.forEach(m => {
    if (!m.desc) return;
    remainingU.forEach(u => {
      if (!u.desc) return;
      const catalogScore = catalogSimilarity_(m.key, u.key);
      const descriptionScore = catalogSimilarity_(m.desc, u.desc);
      const score = 0.65 * catalogScore + 0.35 * descriptionScore;
      if (score >= 40) candidates.push({ m, u, score });
    });
  });
  candidates.sort((a, b) => b.score - a.score || a.m.i - b.m.i || a.u.i - b.u.i);
  [95, 90, 80, 40].forEach(tier => {
    candidates.forEach(({ m, u, score }) => {
      if (score < tier || usedM.has(m.i) || usedU.has(u.i)) return;
      usedM.add(m.i); usedU.add(u.i);
      output[m.i] = [u.price, u.catalog, tier + '%'];
      counts[tier]++;
    });
  });
  const summary = tiers.map(t => t + '%: ' + counts[t]).join(' | ');
  const ui = SpreadsheetApp.getUi();
  const choice = ui.alert('Five-pass match preview',
    summary + '\nUnmatched master rows: ' + (masters.length - usedM.size) +
    '\n\nApply results to Sheet1 columns Z:AB? Existing results will be replaced.',
    ui.ButtonSet.YES_NO);
  if (choice !== ui.Button.YES) return;
  master.getRange(2, 26, mCount, 3).setValues(output);
  ss.toast(summary, 'Five-pass price matching', 10);
}

function catalogSimilarity_(a, b) {
  if (a === b) return 100;
  const longest = Math.max(a.length, b.length);
  if (!longest) return 100;
  let prev = Array.from({ length: b.length + 1 }, (_, i) => i);
  for (let i = 1; i <= a.length; i++) {
    const cur = [i];
    for (let j = 1; j <= b.length; j++) {
      cur[j] = Math.min(cur[j - 1] + 1, prev[j] + 1,
        prev[j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
    }
    prev = cur;
  }
  return 100 * (1 - prev[b.length] / longest);
}

function normaliseDescription_(description) {
  return String(description || '').trim().toUpperCase().replace(/[^A-Z0-9]+/g, ' ').replace(/\\s+/g, ' ').trim();
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
