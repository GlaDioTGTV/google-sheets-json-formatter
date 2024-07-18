function fetchAndRunScript() {
    var url = 'https://raw.githubusercontent.com/GlaDioTGTV/google-sheets-json-formatter/main/google-sheets-json-formatter.js';
    var response = UrlFetchApp.fetch(url);
    var code = response.getContentText();
    
    // Log the fetched code to ensure it's being fetched correctly
    Logger.log(code);
    
    // Execute the fetched code
    eval(code);
  }
  
  function onOpen() {
    fetchAndRunScript(); // Fetch and execute the script from GitHub
    if (typeof onOpenImpl === 'function') {
      onOpenImpl(); // Call the onOpen function from the fetched script, if it exists
    }
  }
  
  function formatJSONInPlace() {
    fetchAndRunScript(); // Fetch and execute the script from GitHub
    if (typeof formatJSONInPlaceImpl === 'function') {
      formatJSONInPlaceImpl(); // Call the formatJSONInPlace function from the fetched script, if it exists
    }
  }
  
  function parseFormattedJSONToColumns() {
    fetchAndRunScript(); // Fetch and execute the script from GitHub
    if (typeof parseFormattedJSONToColumnsImpl === 'function') {
      parseFormattedJSONToColumnsImpl(); // Call the parseFormattedJSONToColumns function from the fetched script, if it exists
    }
  }
  
  function showHelp() {
    fetchAndRunScript(); // Fetch and execute the script from GitHub
    if (typeof showHelpImpl === 'function') {
      showHelpImpl(); // Call the showHelp function from the fetched script, if it exists
    }
  }
  