/**
 * 📊 Google Sheets Data Automator
 * A custom Apps Script suite for parsing, deduplicating, and formatting data sets.
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🚀 Developer Tools')
    // Phase 1: Data Preparation & Cleaning
    .addItem('0. ✨ Prep MASTER (Freeze & Bold)', 'prepMasterSheet')
    .addItem('1. 🧹 Deduplicate Page Paths', 'deduplicatePagePaths')
    .addSeparator()
    
    // Phase 2: Visual Formatting
    .addItem('2. Syntax Highlight URLs', 'applyUrlSyntax')
    .addSeparator()
    
    // Phase 3: State Management & Syncing
    .addItem('3A. 📤 Push Colors: MASTER ➔ All Tabs', 'pushMasterToAll')
    .addItem('3B. 📥 Pull Highlights: All Tabs ➔ MASTER', 'pullAllToMaster')
    .addItem('3C. 🔄 Sync Exact States: All Tabs ➔ MASTER', 'syncExactStatesToMaster')
    .addSeparator()
    
    // Phase 4: Review & Extraction
    .addItem('4. 📋 Create/Update Color Legend', 'createLegendTable')
    .addItem('5. 📝 Extract Highlighted from MASTER', 'extractHighlightedRows')
    .addSeparator()
    
    // Phase 5: Distribution
    .addItem('6. 📂 Split MASTER by Categories (Col D)', 'splitMasterByCategory')
    .addToUi();
}

/**
 * Tool 0: Freezes top row, bolds headers, and sets custom pixel widths for all columns.
 */
function prepMasterSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('MASTER');
  
  if (!sheet) return SpreadsheetApp.getUi().alert('Error: MASTER sheet not found.');

  sheet.setFrozenRows(1);

  // ====================================================================
  // ⚙️ CONFIGURATION AREA: Edit your column widths (in pixels) here!
  // ====================================================================
  const columnWidths = {
    1: 250, // Column A: Data_ID
    2: 400, // Column B: Page path
    3: 100, // Column C: Total users
    4: 180, // Column D: Primary_Category
    5: 220, // Column E: Sub_Category
    6: 180, // Column F: Page_Type_Normalized
    7: 450  // Column G: Target_URL
  };
  // ====================================================================

  if (sheet.getMaxColumns() < 7) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), 7 - sheet.getMaxColumns());
  }

  for (const col in columnWidths) {
    sheet.setColumnWidth(parseInt(col), columnWidths[col]);
  }
  
  sheet.getRange(1, 1, 1, sheet.getMaxColumns()).setFontWeight("bold");
  
  SpreadsheetApp.getUi().alert('MASTER Prepared: Row 1 frozen, headers bolded, and columns resized.');
}

/**
 * Tool 1: Optimized Deduplicator
 * Combines paths, sums metrics, and retains the exact formatting
 * of the row with the highest initial volume.
 */
function deduplicatePagePaths() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('MASTER');
  if (!sheet) return SpreadsheetApp.getUi().alert('Error: MASTER not found.');

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2) return;

  const fullRange = sheet.getRange(1, 1, lastRow, lastCol);
  const values = fullRange.getValues();
  const bgs = fullRange.getBackgrounds();
  const fcs = fullRange.getFontColors();
  const fws = fullRange.getFontWeights();

  const headers = { val: values[0], bg: bgs[0], fc: fcs[0], fw: fws[0] };
  const pathIdx = 1;  // Col B
  const metricIdx = 2; // Col C

  const combinedMap = {};

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    let path = row[pathIdx];
    
    if (typeof path === 'string' && path.endsWith('.html.html')) {
      path = path.replace(/\.html\.html$/, '.html');
      row[pathIdx] = path; 
      
      if (row[6] && typeof row[6] === 'string') {
        row[6] = row[6].replace(/\.html\.html$/, '.html');
      }
    }

    const metricValue = Number(row[metricIdx]) || 0;

    if (!combinedMap[path]) {
      combinedMap[path] = {
        sum: metricValue, maxVal: metricValue,
        bestRow: [...row], bestBg: [...bgs[i]], bestFc: [...fcs[i]], bestFw: [...fws[i]]
      };
    } else {
      combinedMap[path].sum += metricValue;
      if (metricValue > combinedMap[path].maxVal) {
        combinedMap[path].maxVal = metricValue;
        combinedMap[path].bestRow = [...row];
        combinedMap[path].bestBg = [...bgs[i]];
        combinedMap[path].bestFc = [...fcs[i]];
        combinedMap[path].bestFw = [...fws[i]];
      }
    }
  }

  const finalValues = [headers.val];
  const finalBgs = [headers.bg];
  const finalFcs = [headers.fc];
  const finalFws = [headers.fw];

  for (let path in combinedMap) {
    const entry = combinedMap[path];
    entry.bestRow[metricIdx] = entry.sum; 
    finalValues.push(entry.bestRow);
    finalBgs.push(entry.bestBg);
    finalFcs.push(entry.bestFc);
    finalFws.push(entry.bestFw);
  }

  sheet.clear();
  const targetRange = sheet.getRange(1, 1, finalValues.length, lastCol);
  targetRange.setValues(finalValues);
  targetRange.setBackgrounds(finalBgs);
  targetRange.setFontColors(finalFcs);
  targetRange.setFontWeights(finalFws);
  
  const duplicatesCombined = values.length - finalValues.length;
  
  SpreadsheetApp.getUi().alert(`Deduplication Complete: Retained formatting and combined ${duplicatesCombined} duplicate rows.`);
}

/**
 * Tool 2: Syntax Highlights the URLs for easier readability.
 */
function applyUrlSyntax() {
  const range = SpreadsheetApp.getActiveRange();
  const values = range.getValues();
  const richTextValues = [];
  let urlCount = 0;

  const COL_FOLDER = "#0010EE"; // Blue
  const COL_SLASH  = "#D73A49"; // Red
  const COL_FILE   = "#A31515"; // Dark Red (Bold)
  const COL_LOCALE = "#666666"; // Grey

  range.setFontWeight(null).setFontColor(null).setFontStyle(null);

  for (let i = 0; i < values.length; i++) {
    let row = [];
    for (let j = 0; j < values[i].length; j++) {
      let text = values[i][j].toString();
      
      if (!text || text.indexOf('/') === -1) {
        row.push(SpreadsheetApp.newRichTextValue().setText(text).build());
        continue;
      }

      urlCount++;
      let builder = SpreadsheetApp.newRichTextValue().setText(text);
      let len = text.length;

      try {
        builder.setTextStyle(0, len, SpreadsheetApp.newTextStyle().setForegroundColor(COL_FOLDER).build());

        if (text.startsWith("/")) {
          let secondSlash = text.indexOf("/", 1);
          if (secondSlash > 1) {
            builder.setTextStyle(1, secondSlash, SpreadsheetApp.newTextStyle().setForegroundColor(COL_LOCALE).setItalic(true).build());
          }
        }

        for (let k = 0; k < len; k++) {
          if (text[k] === "/") {
            builder.setTextStyle(k, k + 1, SpreadsheetApp.newTextStyle().setForegroundColor(COL_SLASH).setBold(true).build());
          }
        }

        let lastSlash = text.lastIndexOf("/");
        if (lastSlash > -1 && lastSlash < len - 1) {
          builder.setTextStyle(lastSlash + 1, len,
            SpreadsheetApp.newTextStyle().setForegroundColor(COL_FILE).setBold(true).build());
        }

        row.push(builder.build());
      } catch (e) {
        row.push(SpreadsheetApp.newRichTextValue().setText(text).build());
      }
    }
    richTextValues.push(row);
  }

  if (urlCount === 0) {
    SpreadsheetApp.getUi().alert("No URLs found! Make sure you highlight the 'Page path' column with your mouse before clicking the button.");
    return;
  }

  range.setRichTextValues(richTextValues);
  SpreadsheetApp.getActiveSpreadsheet().toast("Successfully styled " + urlCount + " URLs!", "Success ✅");
}

/**
 * Tool 3A: 📤 Push Colors: MASTER ➔ All Tabs
 */
function pushMasterToAll() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName('MASTER');
  if (!masterSheet) return SpreadsheetApp.getUi().alert('Error: MASTER sheet not found.');

  const mData = masterSheet.getDataRange().getValues();
  const mBgs = masterSheet.getDataRange().getBackgrounds();
  const keyIndex = mData[0].indexOf("Page path");

  if (keyIndex === -1) return SpreadsheetApp.getUi().alert('Could not find "Page path" column in MASTER.');

  const masterColorMap = {};
  for (let i = 1; i < mData.length; i++) {
    const path = mData[i][keyIndex];
    if (path) masterColorMap[path] = mBgs[i];
  }

  const skipSheets = ['MASTER', '🎨 Color Legend', '📝 Highlighted Review'];
  let rowsUpdated = 0;

  ss.getSheets().forEach(sheet => {
    if (skipSheets.includes(sheet.getName())) return;

    const data = sheet.getDataRange().getValues();
    const bgs = sheet.getDataRange().getBackgrounds();
    const sKeyIndex = data[0].indexOf("Page path");
    
    if (sKeyIndex === -1) return; 
    
    let sheetChanged = false;

    for (let i = 1; i < data.length; i++) {
      const path = data[i][sKeyIndex];
      const masterColors = masterColorMap[path];

      if (masterColors) {
        const limit = Math.min(bgs[i].length, masterColors.length);
        let rowChanged = false;
        
        for (let c = 0; c < limit; c++) {
          if (bgs[i][c] !== masterColors[c]) {
            bgs[i][c] = masterColors[c];
            rowChanged = true;
          }
        }
        
        if (rowChanged) {
          sheetChanged = true;
          rowsUpdated++;
        }
      }
    }

    if (sheetChanged) {
      sheet.getDataRange().setBackgrounds(bgs);
    }
  });

  ss.toast(`Successfully pushed colors to ${rowsUpdated} rows across all tabs.`, "Success 📤");
}

/**
 * Tool 3B: 📥 Pull Highlights: All Tabs ➔ MASTER
 */
function pullAllToMaster() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName('MASTER');
  if (!masterSheet) return SpreadsheetApp.getUi().alert('Error: MASTER sheet not found.');

  const skipSheets = ['MASTER', '🎨 Color Legend', '📝 Highlighted Review'];
  const childColorMap = {};
  let highlightedRowsFound = 0;

  ss.getSheets().forEach(sheet => {
    if (skipSheets.includes(sheet.getName())) return;

    const data = sheet.getDataRange().getValues();
    const bgs = sheet.getDataRange().getBackgrounds();
    const sKeyIndex = data[0].indexOf("Page path");
    
    if (sKeyIndex === -1) return;

    for (let i = 1; i < data.length; i++) {
      const path = data[i][sKeyIndex];
      if (path && bgs[i][sKeyIndex] !== "#ffffff") {
        childColorMap[path] = bgs[i];
        highlightedRowsFound++;
      }
    }
  });

  if (highlightedRowsFound === 0) {
    return ss.toast("No highlighted rows found in categorized tabs to pull.", "Info ℹ️");
  }

  const mData = masterSheet.getDataRange().getValues();
  const mBgs = masterSheet.getDataRange().getBackgrounds();
  const keyIndex = mData[0].indexOf("Page path");
  let masterChanged = false;

  for (let i = 1; i < mData.length; i++) {
    const path = mData[i][keyIndex];
    const incomingColors = childColorMap[path];

    if (incomingColors) {
      const limit = Math.min(mBgs[i].length, incomingColors.length);
      for (let c = 0; c < limit; c++) {
        if (mBgs[i][c] !== incomingColors[c]) {
          mBgs[i][c] = incomingColors[c];
          masterChanged = true;
        }
      }
    }
  }

  if (masterChanged) {
    masterSheet.getDataRange().setBackgrounds(mBgs);
    ss.toast("Successfully pulled highlighted row colors back into MASTER.", "Success 📥");
  } else {
    ss.toast("MASTER is already up to date with child tabs.", "Up to Date ✅");
  }
}

/**
 * Tool 3C: 🔄 Sync Exact States: All Tabs ➔ MASTER
 */
function syncExactStatesToMaster() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName('MASTER');
  if (!masterSheet) return SpreadsheetApp.getUi().alert('Error: MASTER sheet not found.');

  const skipSheets = ['MASTER', '🎨 Color Legend', '📝 Highlighted Review'];
  const childColorMap = {};
  let rowsScanned = 0;

  ss.getSheets().forEach(sheet => {
    if (skipSheets.includes(sheet.getName())) return;

    const data = sheet.getDataRange().getValues();
    const bgs = sheet.getDataRange().getBackgrounds();
    const sKeyIndex = data[0].indexOf("Page path");
    
    if (sKeyIndex === -1) return;

    for (let i = 1; i < data.length; i++) {
      const path = data[i][sKeyIndex];
      if (path) {
        childColorMap[path] = bgs[i]; 
        rowsScanned++;
      }
    }
  });

  if (rowsScanned === 0) return ss.toast("No valid rows found in categorized tabs to sync.", "Info ℹ️");

  const mData = masterSheet.getDataRange().getValues();
  const mBgs = masterSheet.getDataRange().getBackgrounds();
  const keyIndex = mData[0].indexOf("Page path");
  let masterChanged = false;

  for (let i = 1; i < mData.length; i++) {
    const path = mData[i][keyIndex];
    const incomingColors = childColorMap[path];

    if (incomingColors) {
      const limit = Math.min(mBgs[i].length, incomingColors.length);
      for (let c = 0; c < limit; c++) {
        if (mBgs[i][c] !== incomingColors[c]) {
          mBgs[i][c] = incomingColors[c];
          masterChanged = true;
        }
      }
    }
  }

  if (masterChanged) {
    masterSheet.getDataRange().setBackgrounds(mBgs);
    ss.toast("Successfully synced exact color states back into MASTER.", "Success 🔄");
  } else {
    ss.toast("MASTER is already in perfect sync with child tabs.", "Up to Date ✅");
  }
}

/**
 * Tool 4: Creates or Updates a Legend tab based on colors used, PRESERVING descriptions.
 */
function createLegendTable() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const LEGEND_TAB_NAME = "🎨 Color Legend";
  let legendSheet = ss.getSheetByName(LEGEND_TAB_NAME);
  
  const savedDescriptions = {};
  
  if (legendSheet) {
    const lastRow = legendSheet.getLastRow();
    if (lastRow > 1) {
      const colors = legendSheet.getRange(2, 1, lastRow - 1, 1).getBackgrounds();
      const descriptions = legendSheet.getRange(2, 2, lastRow - 1, 1).getValues();
      
      for (let i = 0; i < colors.length; i++) {
        const hex = colors[i][0];
        const desc = descriptions[i][0];
        if (hex !== "#ffffff" && desc && desc !== "Type description here...") {
          savedDescriptions[hex] = desc;
        }
      }
    }
  } else {
    legendSheet = ss.insertSheet(LEGEND_TAB_NAME);
  }

  const KEY_COLUMN_NAME = "Page path";
  const uniqueColors = new Set();
  
  ss.getSheets().forEach(sheet => {
    const name = sheet.getName();
    if (name === LEGEND_TAB_NAME || name === "📝 Highlighted Review") return;
    
    const data = sheet.getDataRange().getValues();
    const bgs = sheet.getDataRange().getBackgrounds();
    const keyIndex = data[0].indexOf(KEY_COLUMN_NAME);
    
    if (keyIndex !== -1) {
      for (let i = 1; i < bgs.length; i++) {
        if (bgs[i][keyIndex] !== "#ffffff") {
          uniqueColors.add(bgs[i][keyIndex]); 
        }
      }
    }
  });

  legendSheet.clear(); 
  legendSheet.getRange("A1:B1").setValues([["Color Sample", "Meaning / Description"]])
             .setFontWeight("bold").setBackground("#eeeeee");
  
  const colorArray = Array.from(uniqueColors);
  if (colorArray.length === 0) return ss.toast("No colors found!", "Info");

  for (let i = 0; i < colorArray.length; i++) {
    const row = i + 2;
    const colorHex = colorArray[i];
    
    const descriptionText = savedDescriptions[colorHex] ? savedDescriptions[colorHex] : "Type description here...";
    
    legendSheet.getRange(row, 1).setBackground(colorHex);
    legendSheet.getRange(row, 2).setValue(descriptionText);
  }
  
  legendSheet.setColumnWidth(1, 150);
  legendSheet.setColumnWidth(2, 400);
  ss.setActiveSheet(legendSheet);
  ss.toast("Legend updated successfully!", "Success 📋");
}

/**
 * Tool 5: Pulls all colored rows from the MASTER tab into a separate review tab.
 */
function extractHighlightedRows() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let masterSheet = ss.getSheets().find(s => s.getName().toUpperCase().includes("MASTER"));

  if (!masterSheet) return SpreadsheetApp.getUi().alert("Could not find a tab with 'MASTER' in its name.");

  const data = masterSheet.getDataRange().getValues();
  const backgrounds = masterSheet.getDataRange().getBackgrounds();
  const keyIndex = data[0].indexOf("Page path");

  if (keyIndex === -1) return SpreadsheetApp.getUi().alert("Could not find the 'Page path' column in your MASTER tab.");

  const extractedData = [data[0]];
  const extractedBGs = [backgrounds[0]];

  for (let i = 1; i < data.length; i++) {
    if (backgrounds[i][keyIndex] !== "#ffffff") {
      extractedData.push(data[i]);
      extractedBGs.push(backgrounds[i]);
    }
  }

  if (extractedData.length === 1) return ss.toast("No highlighted rows found in the MASTER tab.", "Notice");

  const REVIEW_TAB = "📝 Highlighted Review";
  let reviewSheet = ss.getSheetByName(REVIEW_TAB);
  if (!reviewSheet) {
    reviewSheet = ss.insertSheet(REVIEW_TAB);
  } else {
    reviewSheet.clear();
  }

  const targetRange = reviewSheet.getRange(1, 1, extractedData.length, extractedData[0].length);
  targetRange.setValues(extractedData);
  targetRange.setBackgrounds(extractedBGs);

  ss.setActiveSheet(reviewSheet);
  ss.toast("Successfully extracted " + (extractedData.length - 1) + " rows for review!", "Success 📥");
}

/**
 * Tool 6: Optimized Splitter
 */
function splitMasterByCategory() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName('MASTER');
  if (!masterSheet) return SpreadsheetApp.getUi().alert('Error: MASTER not found.');

  const lastCol = masterSheet.getLastColumn();
  const lastRow = masterSheet.getLastRow();
  if (lastRow < 2) return;

  const fullRange = masterSheet.getRange(1, 1, lastRow, lastCol);
  const values = fullRange.getValues();
  const bgs = fullRange.getBackgrounds();
  const fcs = fullRange.getFontColors();
  const fws = fullRange.getFontWeights();
  
  const masterColWidths = [];
  for (let c = 1; c <= lastCol; c++) {
    masterColWidths.push(masterSheet.getColumnWidth(c));
  }

  const headers = { val: values[0], bg: bgs[0], fc: fcs[0], fw: fws[0] };
  const categoryColIndex = 3; // Column D

  const categoryMap = {};
  for (let i = 1; i < values.length; i++) {
    const cat = values[i][categoryColIndex];
    if (cat) {
      if (!categoryMap[cat]) categoryMap[cat] = [];
      categoryMap[cat].push({
        val: values[i], bg: bgs[i], fc: fcs[i], fw: fws[i]
      });
    }
  }

  Object.keys(categoryMap).forEach(category => {
    let sheet = ss.getSheetByName(category) || ss.insertSheet(category);
    sheet.clear();
    if (sheet.getFilter()) sheet.getFilter().remove();

    const catRows = categoryMap[category];
    const numRows = catRows.length + 1; 

    if (sheet.getMaxRows() < numRows) {
      sheet.insertRowsAfter(sheet.getMaxRows(), numRows - sheet.getMaxRows() + 5);
    }
    if (sheet.getMaxColumns() < lastCol) {
      sheet.insertColumnsAfter(sheet.getMaxColumns(), lastCol - sheet.getMaxColumns());
    }

    for (let c = 1; c <= lastCol; c++) {
      sheet.setColumnWidth(c, masterColWidths[c - 1]);
    }

    const writeValues = [headers.val];
    const writeBgs = [headers.bg];
    const writeFcs = [headers.fc];
    const writeFws = [headers.fw];

    catRows.forEach(row => {
      writeValues.push(row.val);
      writeBgs.push(row.bg);
      writeFcs.push(row.fc);
      writeFws.push(row.fw);
    });

    const targetRange = sheet.getRange(1, 1, numRows, lastCol);
    targetRange.setValues(writeValues);
    targetRange.setBackgrounds(writeBgs);
    targetRange.setFontColors(writeFcs);
    targetRange.setFontWeights(writeFws);

    sheet.setFrozenRows(1);

    if (numRows > 1) {
      targetRange.createFilter();
    }
  });

  SpreadsheetApp.getUi().alert('Success: Split complete. MASTER column widths retained!');
}
