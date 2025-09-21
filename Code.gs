function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Kruti Dev Tools')
      .addItem('Open Helper Sidebar', 'showSidebar')
      .addToUi();
}

function showSidebar() {
  var htmlOutput = HtmlService.createHtmlOutputFromFile('Sidebar')
      .setTitle('Kruti Dev & Hindi Helper');
  SpreadsheetApp.getUi().showSidebar(htmlOutput);
}

function krutiDevToUnicode(krutiDevText) {
  try {
    // Call the robust conversion function from KrutiDevConversionLibrary.gs
    return convertKrutiDevToUnicode(krutiDevText);
  } catch (e) {
    Logger.log('Error in krutiDevToUnicode: ' + e.message + ' Input: ' + krutiDevText);
    return 'Conversion Error (Kruti Dev to Unicode): ' + krutiDevText;
  }
}

function unicodeToKrutiDev(unicodeText) {
  try {
    return convertUnicodeToKrutiDev(unicodeText);
  } catch (e) {
    Logger.log('Error in unicodeToKrutiDev: ' + e.message + ' Input: ' + unicodeText);
    return 'Conversion Error (Unicode to Kruti Dev): ' + unicodeText;
  }
}

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
        row.push(values[i][j]);
      }
    }
    newValues.push(row);
  }

  range.setValues(newValues);
  return 'Selected cells converted to Kruti Dev.';
}

function downloadAsKrutiDevCSV() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var range = sheet.getDataRange();
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

  var blob = Utilities.newBlob(csvFile, 'text/csv;charset=utf-8;', fileName); 

  return blob.getDataAsString();
}

function printSheet() {
  
  var htmlOutput = HtmlService.createHtmlOutput('<script>window.print(); google.script.host.close();</script>')
      .setWidth(1) 
      .setHeight(1);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Print Sheet');
  return "Print dialog opened.";
}
