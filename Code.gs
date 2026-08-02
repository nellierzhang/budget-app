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

      const data = sheet.getRange(1, 1, lastRow, Math.max(lastCol, 6)).getValues();
      const backgrounds = sheet.getRange(1, 1, lastRow, Math.max(lastCol, 6)).getBackgrounds();
      const fontColors = sheet.getRange(1, 1, lastRow, Math.max(lastCol, 6)).getFontColors();

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
      rows.forEach((tx, i) => {
        const rowNum = appendAt + i;
        sheet.getRange(rowNum, 1).setValue(tx.date);
        sheet.getRange(rowNum, 2).setValue(tx.business);
        sheet.getRange(rowNum, 3).setValue(tx.amount);
        sheet.getRange(rowNum, 4).setValue(tx.category);
        sheet.getRange(rowNum, 5).setValue(tx.spending_category);
        sheet.getRange(rowNum, 6).setValue(tx.notes || '');

        // Amazon transactions get a bright yellow highlight to flag for review,
        // except Amazon Grocery purchases which are already well-categorized.
        // Color is applied only to the category/spending_category columns (D–E)
        // so date, business, and amount remain uncolored.
        const isAmazon = String(tx.business).toUpperCase().includes('AMAZON');
        const isAmazonGrocery = isAmazon && String(tx.category).trim() === 'Grocery';
        if (isAmazon && !isAmazonGrocery) {
          sheet.getRange(rowNum, 4, 1, 2).setBackground('#fff2cc');
          sheet.getRange(rowNum, 4, 1, 2).setFontColor('#7d5a00');
        } else {
          // Apply background color matching the category's existing color in the sheet
          const color = categoryColors[String(tx.category).trim()];
          if (color) {
            sheet.getRange(rowNum, 4, 1, 2).setBackground(color.bg);
            if (color.text) sheet.getRange(rowNum, 4, 1, 2).setFontColor(color.text);
          }
        }
      });
      results.push(`${month}: wrote ${rows.length} row(s) starting at row ${appendAt}`);
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
    const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    const backgrounds = sheet.getRange(2, 1, lastRow - 1, lastCol).getBackgrounds();
    const fontColors = sheet.getRange(2, 1, lastRow - 1, lastCol).getFontColors();
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
