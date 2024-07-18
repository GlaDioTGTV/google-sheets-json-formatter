(function(global) {
  function onOpenImpl() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('JSON Formatter')
        .addItem('Format JSON In Place', 'formatJSONInPlace')
        .addItem('Parse Formatted JSON to Columns', 'parseFormattedJSONToColumns')
        .addSeparator()
        .addItem('Help', 'showHelp')
        .addToUi();
  }

  function formatJSONInPlaceImpl() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var range = sheet.getActiveRange();
    var values = range.getValues();
    
    for (var i = 1; i < values.length; i++) { // Start from the second row
      for (var j = 0; j < values[i].length; j++) {
        var jsonString = values[i][j];
        try {
          var jsonData = JSON.parse(jsonString);
          var keys = Object.keys(jsonData);
          var formatted = keys.map(function(key) {
            return key + ": " + jsonData[key];
          }).join(", ");
          values[i][j] = formatted;
        } catch (e) {
          values[i][j] = "Invalid JSON";
        }
      }
    }
    
    range.setValues(values);
  }

  function parseFormattedJSONToColumnsImpl() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var range = sheet.getActiveRange();
    var values = range.getValues();
    
    var allKeys = new Set();
    var formattedObjects = [];
    
    for (var i = 1; i < values.length; i++) { // Start from the second row
      for (var j = 0; j < values[i].length; j++) {
        var formattedString = values[i][j];
        if (!formattedString || formattedString.trim() === "") {
          continue; // Skip empty cells
        }
        var pairs = formattedString.split(', ');
        var jsonObject = {};
        for (var k = 0; k < pairs.length; k++) {
          var pair = pairs[k].split(': ');
          if (pair.length == 2) {
            var key = pair[0];
            var value = pair[1];
            jsonObject[key] = value;
            allKeys.add(key);
          }
        }
        formattedObjects.push(jsonObject);
      }
    }
    
    if (allKeys.size === 0) {
      SpreadsheetApp.getUi().alert("No valid formatted JSON data found in the selected range.");
      return;
    }

    var headers = Array.from(allKeys);
    var output = [headers];
    
    formattedObjects.forEach(function(jsonObject) {
      var row = headers.map(function(header) {
        return jsonObject[header] !== undefined ? jsonObject[header] : null;
      });
      output.push(row);
    });
    
    var startColumn = range.getLastColumn() + 1;
    var outputRange = sheet.getRange(range.getRow(), startColumn, output.length, headers.length);
    outputRange.setValues(output);
  }

  function showHelpImpl() {
    var ui = SpreadsheetApp.getUi();
    var helpText = 'This menu provides the following functionalities:\n\n' +
      '1. Format JSON In Place: This option formats the JSON data in the selected range as "key: value" pairs within the same cells, starting from the second row to exclude the header.\n\n' +
      '2. Parse Formatted JSON to Columns: This option parses the formatted JSON strings in the selected range, extracts the keys, and arranges the data into new columns, while keeping the original column intact.\n\n' +
      'Author: med.maddouri@gmail.com';
    ui.alert('Help - JSON Formatter', helpText, ui.ButtonSet.OK);
  }
  
  global.onOpenImpl = onOpenImpl;
  global.formatJSONInPlaceImpl = formatJSONInPlaceImpl;
  global.parseFormattedJSONToColumnsImpl = parseFormattedJSONToColumnsImpl;
  global.showHelpImpl = showHelpImpl;
})(this);
