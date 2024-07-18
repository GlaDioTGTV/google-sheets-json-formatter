# Google Sheets JSON Formatter

Google Sheets JSON Formatter is a Google Apps Script that adds a custom menu to Google Sheets, enabling users to format JSON data as "key: value" pairs within cells and parse formatted JSON into columns. This tool simplifies working with JSON data in Google Sheets, making it easy to organize and visualize data directly in your spreadsheets.

## Features

- **Format JSON In Place**: Convert JSON strings into readable "key: value" pairs within the selected range.
- **Parse Formatted JSON to Columns**: Extract keys from formatted JSON and arrange the data into new columns, preserving the original format.

## How to Use

1. Open your Google Sheets document.
2. Go to `Extensions > Apps Script`.
3. Delete any default code in the script editor and paste the provided code from `Code.js`.
4. Save the script and reload the Google Sheet.
5. Use the `JSON Formatter` menu in the Google Sheets UI.

### Example

#### Formatting JSON in Place

The formatter is used when the JSON data is in multiple lines within a single cell. Here's how to use it:

1. Enter your JSON data into a cell:
   ```json
   {
     "name": "Sarah",
     "age": 30,
     "city": "Tunis"
   }
   ```
2. Select the cell with the JSON data.
3. From the `JSON Formatter` menu, choose `Format JSON In Place` to convert the JSON string into readable "key: value" pairs within the same cell:
   ```plaintext
   name: Sarah, age: 30, city: Tunis
   ```

#### Parsing Formatted JSON to Columns

The parser is used to extract keys from formatted JSON strings and arrange the data into columns:

1. Ensure your JSON data is formatted in place as "key: value" pairs within the cell:
   ```plaintext
   name: Sarah, age: 30, city: Tunis
   ```
2. Select the cell with the formatted JSON data.
3. From the `JSON Formatter` menu, choose `Parse Formatted JSON to Columns`. The data will be extracted and arranged into new columns, starting from the first empty column to the right of the selected cell.

## Installation

1. Open your Google Sheets document.
2. Navigate to `Extensions > Apps Script`.
3. Delete any default code in the script editor and paste in the provided code from `google-sheets-json-formatter.js`.
4. Save the project with an appropriate name.

## Contribution

Feel free to fork this repository, make improvements, and submit pull requests. Your contributions are welcome!

## License

This project is licensed under the MIT License - see the LICENSE.md file for details.

## Author

- Mohamed Maddouri - med.maddouri@gmail.com