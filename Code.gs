/**
 * @OnlyCurrentDoc
 * The above line ensures that the script can only access the current document.
 */

// This function runs automatically when the spreadsheet is opened.
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Add a custom menu named "Kruti Dev Tools"
  ui.createMenu('Kruti Dev Tools')
      .addItem('Open Helper Sidebar', 'showSidebar') // Adds a menu item to open the sidebar
      .addToUi();
}

// This function shows the sidebar.
function showSidebar() {
  var htmlOutput = HtmlService.createHtmlOutputFromFile('Sidebar')
      .setTitle('Kruti Dev & Hindi Helper');
  SpreadsheetApp.getUi().showSidebar(htmlOutput);
}

// ============================================================================
// Core Conversion Functions (Now calling the robust library)
// ============================================================================

/**
 * Converts Kruti Dev text to Unicode using the robust conversion library.
 * @param {string} krutiDevText The text in Kruti Dev font.
 * @return {string} The converted text in Unicode.
 */
function krutiDevToUnicode(krutiDevText) {
  try {
    // Call the robust conversion function from KrutiDevConversionLibrary.gs
    return convertKrutiDevToUnicode(krutiDevText);
  } catch (e) {
    Logger.log('Error in krutiDevToUnicode: ' + e.message + ' Input: ' + krutiDevText);
    return 'Conversion Error (Kruti Dev to Unicode): ' + krutiDevText;
  }
}

/**
 * Converts Unicode text to Kruti Dev using the robust conversion library.
 * @param {string} unicodeText The text in Unicode font.
 * @return {string} The converted text in Kruti Dev.
 */
function unicodeToKrutiDev(unicodeText) {
  try {
    // Call the robust conversion function from KrutiDevConversionLibrary.gs
    return convertUnicodeToKrutiDev(unicodeText);
  } catch (e) {
    Logger.log('Error in unicodeToKrutiDev: ' + e.message + ' Input: ' + unicodeText);
    return 'Conversion Error (Unicode to Kruti Dev): ' + unicodeText;
  }
}

// ============================================================================
// Functions to interact with the active spreadsheet and apply conversions
// These remain largely the same, but now call the robust conversion functions.
// ============================================================================

/**
 * Converts selected cells from Kruti Dev to Unicode.
 */
function convertSelectionToUnicode() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var range = sheet.getActiveRange();
  if (!range) {
    return 'No cells selected.';
  }

  var values = range.getValues();
  var newValues = [];

  for (var i = 0; i < values.length; i++) {
    var row = [];
    for (var j = 0; j < values[i].length; j++) {
      if (typeof values[i][j] === 'string') {
        row.push(krutiDevToUnicode(values[i][j]));
      } else {
        row.push(values[i][j]); // Keep non-string values as they are
      }
    }
    newValues.push(row);
  }

  range.setValues(newValues);
  return 'Selected cells converted to Unicode.';
}

/**
 * Converts selected cells from Unicode to Kruti Dev.
 */
function convertSelectionToKrutiDev() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var range = sheet.getActiveRange();
  if (!range) {
    return 'No cells selected.';
  }

  var values = range.getValues();
  var newValues = [];

  for (var i = 0; i < values.length; i++) {
    var row = [];
    for (var j = 0; j < values[i].length; j++) {
      if (typeof values[i][j] === 'string') {
        row.push(unicodeToKrutiDev(values[i][j]));
      } else {
        row.push(values[i][j]); // Keep non-string values as they are
      }
    }
    newValues.push(row);
  }

  range.setValues(newValues);
  return 'Selected cells converted to Kruti Dev.';
}


// ============================================================================
// Functions for Downloading and Printing
// These also remain largely the same, using the now robust conversion functions.
// ============================================================================

/**
 * Creates a temporary Excel file (CSV in this context) with Kruti Dev converted data
 * and provides a download link.
 */
function downloadAsKrutiDevCSV() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var range = sheet.getDataRange(); // Get all data
  var values = range.getValues();

  var csvContent = [];
  for (var i = 0; i < values.length; i++) {
    var row = [];
    for (var j = 0; j < values[i].length; j++) {
      var cellValue = values[i][j];
      if (typeof cellValue === 'string') {
        // Convert to Kruti Dev for download
        row.push('"' + unicodeToKrutiDev(cellValue).replace(/"/g, '""') + '"');
      } else {
        row.push('"' + String(cellValue).replace(/"/g, '""') + '"');
      }
    }
    csvContent.push(row.join(','));
  }

  var fileName = sheet.getName() + '_KrutiDev_' + new Date().getTime() + '.csv';
  var csvFile = csvContent.join('\n');

  // Create a blob for download
  var blob = Utilities.newBlob(csvFile, 'text/csv;charset=utf-8;', fileName); // Specify charset for better compatibility
  // In Apps Script, we can't directly trigger a browser download from server-side.
  // We return the content, and the client-side (Sidebar.html) handles the download.
  return blob.getDataAsString(); // Returns the string content
}


/**
 * Opens the Google Sheets print dialog.
 */
function printSheet() {
  // This function opens a modal dialog with JavaScript that triggers the browser's print function.
  // The user will see the print dialog for the current view of the sheet.
  // For Kruti Dev printing, the sheet content *must* be in Unicode for Google Sheets to render it correctly for printing.
  var htmlOutput = HtmlService.createHtmlOutput('<script>window.print(); google.script.host.close();</script>')
      .setWidth(1) // Make it tiny as it's just a trigger
      .setHeight(1);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Print Sheet');
  return "Print dialog opened."; // Return a message for the client-side
}
