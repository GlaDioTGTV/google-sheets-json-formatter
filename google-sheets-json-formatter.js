function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('JSON Formatter')
      .addItem('Format JSON In Place', 'formatJSONInPlace')
      .addItem('Parse Formatted JSON to Columns', 'parseFormattedJSONToColumns')
      .addSeparatora()
      .addItem('Revert to JSON', 'revertToJSON')
      .addSeparator()
      .addItem('Help', 'showHelp')
      .addToUi();
}

function formatJSONInPlace() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var range = sheet.getActiveRange();
  var values = range.getValues();
  
  for (var i = 1; i < values.length; i++) { // Start from the second row
    for (var j = 0; j < values[i].length; j++) {
      var jsonString = values[i][j];
      try {
        var jsonData = JSON.parse(jsonString);
        var formatted = Object.keys(jsonData).map(function(key) {
          return key + ": " + jsonData[key];
        }).join(", ");
        values[i][j] = formatted;
      } catch (e) {
        values[i][j] = "Invalid JSON"; // Mark invalid JSON strings
      }
    }
  }
  
  range.setValues(values); // Set the modified values back to the range
}

function parseFormattedJSONToColumns() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var range = sheet.getActiveRange();
  var values = range.getValues();
  
  var allKeys = new Set();
  var formattedObjects = [];
  
  // Collect all keys from the formatted strings
  for (var i = 1; i < values.length; i++) { // Start from the second row
    for (var j = 0; j < values[i].length; j++) {
      var formattedString = values[i][j];
      if (!formattedString || formattedString.trim() === "") {
        continue; // Skip empty cells
      }
      var jsonObject = {};
      formattedString.split(', ').forEach(function(pair) {
        var [key, value] = pair.split(': ');
        if (key && value) {
          jsonObject[key] = value;
          allKeys.add(key);
        }
      });
      formattedObjects.push(jsonObject);
    }
  }
  
  if (allKeys.size === 0) {
    SpreadsheetApp.getUi().alert("No valid formatted JSON data found in the selected range.");
    return;
  }

  var headers = Array.from(allKeys);
  var output = [headers];
  
  // Fill in the values for each formatted JSON object
  formattedObjects.forEach(function(jsonObject) {
    var row = headers.map(function(header) {
      return jsonObject[header] !== undefined ? jsonObject[header] : null;
    });
    output.push(row);
  });
  
  // Set the values in the sheet starting from the first empty column
  var startColumn = range.getLastColumn() + 1;
  var outputRange = sheet.getRange(range.getRow(), startColumn, output.length, headers.length);
  outputRange.setValues(output);
}

function revertToJSON() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var range = sheet.getDataRange();
  var values = range.getValues();

  for (var i = 1; i < values.length; i++) { // Start from the second row
    for (var j = 0; j < values[i].length; j++) {
      var formattedString = values[i][j];
      if (formattedString && typeof formattedString === 'string' && formattedString.includes(":")) {
        try {
          var jsonObject = {};
          formattedString.split(',').forEach(function(pair) {
            var [key, value] = pair.split(':');
            if (key && value) {
              jsonObject[key.trim()] = value.trim();
            }
          });
          values[i][j] = JSON.stringify(jsonObject);
        } catch (e) {
          values[i][j] = "Invalid Format"; // Mark invalid formatted strings
        }
      }
    }
  }

  range.setValues(values); // Set the modified values back to the range
}

function showHelp() {
  var ui = SpreadsheetApp.getUi();
  var helpText = 'This menu provides the following functionalities:\n\n' +
    '1. Format JSON In Place: This option formats the JSON data in the selected range as "key: value" pairs within the same cells, starting from the second row to exclude the header.\n\n' +
    '2. Parse Formatted JSON to Columns: This option parses the formatted JSON strings in the selected range, extracts the keys, and arranges the data into new columns, while keeping the original column intact.\n\n' +
    '3. Revert to JSON: This option reverts formatted "key: value" pairs back to JSON starting from the second row.\n\n' +
    'Author: med.maddouri@gmail.com';
  ui.alert('Help - JSON Formatter', helpText, ui.ButtonSet.OK);
}