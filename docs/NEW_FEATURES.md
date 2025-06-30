# Excel Data Validation, Conditional Formatting, and VBA Tools

This document describes the new Excel tools added to the MCP server for data validation, conditional formatting, and VBA functionality.

## Data Validation Tool

### `excel_add_data_validation`

Adds data validation to Excel cells including dropdown lists and various validation types.

**Parameters:**

- `fileAbsolutePath` (string, required): Absolute path to the Excel file
- `sheetName` (string, required): Sheet name in the Excel file
- `cellRange` (string, required): Range of cells to apply data validation (e.g., "A1:A10")
- `validationType` (string, required): Type of validation
  - `list`: Dropdown list validation
  - `whole`: Whole number validation
  - `decimal`: Decimal number validation
  - `date`: Date validation
  - `time`: Time validation
  - `textLength`: Text length validation
  - `custom`: Custom formula validation
- `options` (object, optional): Data validation options

**Options Object:**

- `dropdownList` (array of strings): List of options for dropdown validation
- `formula1` (string): First formula for validation criteria
- `formula2` (string): Second formula for validation criteria (for between/notBetween)
- `operator` (string): Validation operator
  - `between`, `notBetween`, `equal`, `notEqual`
  - `greaterThan`, `lessThan`, `greaterThanOrEqual`, `lessThanOrEqual`
- `showErrorMessage` (boolean): Whether to show error message on invalid input
- `errorTitle` (string): Title for error message
- `errorMessage` (string): Error message text
- `showInputMessage` (boolean): Whether to show input message when cell is selected
- `inputTitle` (string): Title for input message
- `inputMessage` (string): Input message text

**Example - Dropdown List:**

```json
{
  "fileAbsolutePath": "C:\\data\\workbook.xlsx",
  "sheetName": "Sheet1",
  "cellRange": "A1:A10",
  "validationType": "list",
  "options": {
    "dropdownList": ["Option 1", "Option 2", "Option 3"],
    "showInputMessage": true,
    "inputTitle": "Select Option",
    "inputMessage": "Please select an option from the dropdown list",
    "showErrorMessage": true,
    "errorTitle": "Invalid Selection",
    "errorMessage": "Please select a valid option from the list"
  }
}
```

**Example - Number Range:**

```json
{
  "fileAbsolutePath": "C:\\data\\workbook.xlsx",
  "sheetName": "Sheet1",
  "cellRange": "B1:B10",
  "validationType": "whole",
  "options": {
    "operator": "between",
    "formula1": "1",
    "formula2": "100",
    "showInputMessage": true,
    "inputTitle": "Enter Number",
    "inputMessage": "Please enter a number between 1 and 100"
  }
}
```

## Conditional Formatting Tool

### `excel_add_conditional_formatting`

Adds conditional formatting to Excel cells with various formatting types.

**Parameters:**

- `fileAbsolutePath` (string, required): Absolute path to the Excel file
- `sheetName` (string, required): Sheet name in the Excel file
- `cellRange` (string, required): Range of cells to apply conditional formatting (e.g., "A1:A10")
- `conditions` (object, required): Conditional formatting conditions

**Conditions Object:**

- `type` (string): Type of conditional formatting
  - `cellValue`: Cell value based formatting
  - `expression`: Formula/expression based formatting
  - `colorScale`: Color scale formatting
  - `dataBar`: Data bar formatting
  - `iconSet`: Icon set formatting
- `criteria` (string): Criteria for cell value formatting
  - `greaterThan`, `lessThan`, `between`, `equal`, `notEqual`
  - `greaterThanOrEqual`, `lessThanOrEqual`
- `value1` (string): First value for comparison
- `value2` (string): Second value for comparison (for between)
- `formula` (string): Formula for expression-based formatting
- `format` (object): Formatting to apply (font, fill, border)
- `colorScale` (object): Color scale options
- `dataBar` (object): Data bar options
- `iconSet` (object): Icon set options

**Format Object:**

- `font` (object): Font formatting
  - `bold` (boolean): Bold text
  - `italic` (boolean): Italic text
  - `color` (string): Font color (hex format, e.g., "#FF0000")
  - `size` (number): Font size
- `fill` (object): Fill formatting
  - `type` (string): Fill type
  - `color` (array of strings): Fill colors
- `border` (array): Border formatting

**Color Scale Object:**

- `minType` (string): Minimum value type (`num`, `percent`, `percentile`, `formula`, `min`, `max`)
- `minValue` (string): Minimum value
- `minColor` (string): Minimum color (hex format)
- `maxType` (string): Maximum value type
- `maxValue` (string): Maximum value  
- `maxColor` (string): Maximum color (hex format)
- `midType` (string): Middle value type (for 3-color scale)
- `midValue` (string): Middle value
- `midColor` (string): Middle color

**Data Bar Object:**

- `minType` (string): Minimum value type
- `minValue` (string): Minimum value
- `maxType` (string): Maximum value type
- `maxValue` (string): Maximum value
- `color` (string): Bar color (hex format)
- `showValue` (boolean): Whether to show values

**Example - Highlight Greater Than:**

```json
{
  "fileAbsolutePath": "C:\\data\\workbook.xlsx",
  "sheetName": "Sheet1",
  "cellRange": "A1:A10",
  "conditions": {
    "type": "cellValue",
    "criteria": "greaterThan",
    "value1": "50",
    "format": {
      "font": {
        "bold": true,
        "color": "#FFFFFF"
      },
      "fill": {
        "type": "pattern",
        "color": ["#FF0000"]
      }
    }
  }
}
```

**Example - Color Scale:**

```json
{
  "fileAbsolutePath": "C:\\data\\workbook.xlsx",
  "sheetName": "Sheet1",
  "cellRange": "A1:A10",
  "conditions": {
    "type": "colorScale",
    "colorScale": {
      "minType": "min",
      "minColor": "#FF0000",
      "maxType": "max",
      "maxColor": "#00FF00"
    }
  }
}
```

**Example - Data Bars:**

```json
{
  "fileAbsolutePath": "C:\\data\\workbook.xlsx",
  "sheetName": "Sheet1",
  "cellRange": "A1:A10",
  "conditions": {
    "type": "dataBar",
    "dataBar": {
      "minType": "min",
      "maxType": "max",
      "color": "#0000FF",
      "showValue": true
    }
  }
}
```

## VBA Tools

### `excel_execute_vba`

Executes VBA code on an Excel worksheet (Windows OLE backend only).

**Parameters:**

- `fileAbsolutePath` (string, required): Absolute path to the Excel file
- `sheetName` (string, required): Sheet name in the Excel file
- `vbaCode` (string, required): VBA code to execute

**Example:**

```json
{
  "fileAbsolutePath": "C:\\data\\workbook.xlsx",
  "sheetName": "Sheet1",
  "vbaCode": "Range(\"A1\").Value = \"Hello from VBA\""
}
```

### `excel_add_vba_module`

Adds a VBA module to an Excel workbook (Windows OLE backend only).

**Parameters:**

- `fileAbsolutePath` (string, required): Absolute path to the Excel file
- `moduleName` (string, required): Name for the VBA module
- `vbaCode` (string, required): VBA code to add to the module

**Example:**

```json
{
  "fileAbsolutePath": "C:\\data\\workbook.xlsx",
  "moduleName": "MyModule",
  "vbaCode": "Sub HelloWorld()\n    MsgBox \"Hello, World!\"\nEnd Sub\n\nFunction AddNumbers(a As Double, b As Double) As Double\n    AddNumbers = a + b\nEnd Function"
}
```

## Backend Support

| Feature | Excelize Backend | OLE Backend (Windows) |
|---------|:----------------:|:--------------------:|
| Data Validation | ✅ | ✅ |
| Conditional Formatting | ✅ | ✅ |
| VBA Execution | ❌ | ✅ |
| VBA Modules | ❌ | ✅ |

**Notes:**

- VBA functionality is only supported on Windows with the OLE backend
- The excelize backend will return an error for VBA-related operations
- Data validation and conditional formatting work on both backends but may have different feature sets
- Color scale and data bar formatting require Excel 2007 or later when using OLE backend

## Error Handling

All tools return appropriate error messages for:

- Invalid file paths
- Missing sheets
- Invalid cell ranges
- Unsupported operations (e.g., VBA on excelize backend)
- Missing required parameters
- Invalid parameter values

The tools automatically save the Excel file after making changes.
