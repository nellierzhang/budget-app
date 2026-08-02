const MONTH_NAMES = ['January','February','March','April','May','June','July','August','September','October','November','December'];

function doGet(e) {
  if (e.parameter.action === 'read') {
    return readExistingKeys();
  }
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'Budget webhook active' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function readExistingKeys() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const keys = [];
    const merchantHistory = {};
    const categoryColors = {}; // category name → { bg, text } hex colors from sheet

    MONTH_NAMES.forEach(month => {
      const sheet = ss.getSheetByName(month);
      if (!sheet) return;
      const lastRow = sheet.getLastRow();
      const lastCol = sheet.getLastColumn();
      if (lastRow < 2) return;

      // Fetch all three data types with one getRange() reference to avoid
      // creating redundant range objects (each getRange call is cheap, but
      // consolidating makes intent clearer and is slightly more efficient).
      const dataRange = sheet.getRange(1, 1, lastRow, Math.max(lastCol, 6));
      const data = dataRange.getValues();
      const backgrounds = dataRange.getBackgrounds();
      const fontColors = dataRange.getFontColors();

      // Iterate newest → oldest so the first write into merchantHistory wins (= most recent categorization)
      for (let i = data.length - 1; i >= 1; i--) {
        const row = data[i];
        const date = row[0];
        const business = row[1];
        const amount = row[2];
        const cat = row[3];
        const spendingCat = row[4];

        if (!date && !business) continue;

        // Build duplicate-detection key
        let dateStr = '';
        if (date instanceof Date) {
          dateStr = `${date.getMonth()+1}/${date.getDate()}/${String(date.getFullYear()).slice(-2)}`;
        } else {
          dateStr = String(date);
        }
        const key = `${dateStr}|${String(business).toUpperCase().trim()}|${Math.abs(parseFloat(amount)||0).toFixed(2)}`;
        keys.push(key);

        // Build merchant history map — only record if we have a category and haven't seen this merchant yet
        const merchantKey = String(business).toUpperCase().trim();
        if (merchantKey && cat && !merchantHistory[merchantKey]) {
          merchantHistory[merchantKey] = {
            category: String(cat),
            spending_category: String(spendingCat || '')
          };
        }

        // Capture row background color keyed by category (col 4 = index 3)
        // Only store if the row has a non-white, non-null background
        const catKey = String(cat).trim();
        if (catKey && !categoryColors[catKey]) {
          const bg = backgrounds[i][3]; // category column background
          const fg = fontColors[i][3];
          if (bg && bg !== '#ffffff' && bg !== '#000000' && bg !== 'white') {
            categoryColors[catKey] = { bg, text: fg && fg !== '#000000' ? fg : null };
          }
        }
      }
    });

    return ContentService
      .createTextOutput(JSON.stringify({ success: true, existingKeys: keys, merchantHistory, categoryColors }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const transactions = data.transactions;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const results = [];

    // Build a category→color map from existing sheet rows so new rows get the same colors
    const categoryColors = buildCategoryColorMap(ss);

    const byMonth = {};
    transactions.forEach(tx => {
      if (!byMonth[tx.month]) byMonth[tx.month] = [];
      byMonth[tx.month].push(tx);
    });

    Object.entries(byMonth).forEach(([month, rows]) => {
      const sheet = ss.getSheetByName(month);
      if (!sheet) { results.push(`Tab not found: ${month}`); return; }
      const lastRow = getLastDataRow(sheet);
      const appendAt = lastRow + 1;
      const n = rows.length;

      // Build value matrix and color arrays in one pass, then write with a
      // single setValues() + two setBackgrounds()/setFontColors() calls instead
      // of 6+ individual setValue/setBackground calls per row.
      const values = [];
      const backgrounds = [];
      const fontColors = [];

      rows.forEach(tx => {
        values.push([tx.date, tx.business, tx.amount, tx.category, tx.spending_category, tx.notes || '']);

        // Amazon transactions get a bright yellow highlight to flag for review,
        // except Amazon Grocery purchases which are already well-categorized.
        // Color is applied only to the category/spending_category columns (D–E)
        // so date, business, and amount remain uncolored.
        const isAmazon = String(tx.business).toUpperCase().includes('AMAZON');
        const isAmazonGrocery = isAmazon && String(tx.category).trim() === 'Grocery';

        let catBg = null;
        const isAmazonHighlight = isAmazon && !isAmazonGrocery;
        if (isAmazonHighlight) {
          catBg = '#fff2cc';
        } else {
          const color = categoryColors[String(tx.category).trim()];
          if (color) catBg = color.bg;
        }

        // backgrounds: highlight spans all 6 columns so the full row is visually marked;
        // fontColors: explicitly set every column to black so inherited row formatting is always cleared;
        // Amazon D–E get dark brown for contrast against the yellow background
        const deFontColor = isAmazonHighlight ? '#7d5a00' : '#000000';
        backgrounds.push([catBg, catBg, catBg, catBg, catBg, catBg]);
        fontColors.push(['#000000', '#000000', '#000000', deFontColor, deFontColor, '#000000']);
      });

      const range = sheet.getRange(appendAt, 1, n, 6);
      range.setValues(values);
      range.setBackgrounds(backgrounds);
      range.setFontColors(fontColors);

      results.push(`${month}: wrote ${n} row(s) starting at row ${appendAt}`);
    });

    return ContentService
      .createTextOutput(JSON.stringify({ success: true, results }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function buildCategoryColorMap(ss) {
  const categoryColors = {};
  MONTH_NAMES.forEach(month => {
    const sheet = ss.getSheetByName(month);
    if (!sheet) return;
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;
    const lastCol = Math.max(sheet.getLastColumn(), 6);
    const dataRange = sheet.getRange(2, 1, lastRow - 1, lastCol);
    const data = dataRange.getValues();
    const backgrounds = dataRange.getBackgrounds();
    const fontColors = dataRange.getFontColors();
    for (let i = data.length - 1; i >= 0; i--) {
      const cat = String(data[i][3]).trim();
      if (!cat || categoryColors[cat]) continue;
      const bg = backgrounds[i][3];
      const fg = fontColors[i][3];
      if (bg && bg !== '#ffffff' && bg !== '#000000' && bg !== 'white') {
        categoryColors[cat] = { bg, text: fg && fg !== '#000000' ? fg : null };
      }
    }
  });
  return categoryColors;
}

function getLastDataRow(sheet) {
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][0] || data[i][1] || data[i][2]) return i + 1;
  }
  return 1;
}
