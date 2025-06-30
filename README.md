# Excel MCP Server

![Excel MCP Server Icon](docs/img/icon-800.png)

[![NPM Version](https://img.shields.io/npm/v/ms-excel-mcp-server)](https://www.npmjs.com/package/ms-excel-mcp-server)
[![smithery badge](https://smithery.ai/badge/ms-excel-mcp-server)](https://smithery.ai/server/ms-excel-mcp-server)

A Model Context Protocol (MCP) server that reads and writes MS Excel data.

## Features

- Read/Write text values
- Read/Write formulas
- Create new sheets
- **✨ NEW: Data validation with dropdown lists**
- **✨ NEW: Conditional formatting (highlighting, color scales, data bars)**
- **✨ NEW: VBA code execution and module creation**

**🪟Windows only:**

- Live editing
- Capture screen image from a sheet
- **✨ VBA functionality (OLE backend)**

For more details, see the [tools](#tools) section.

## Requirements

- Node.js 20.x or later

## Supported file formats

- xlsx (Excel book)
- xlsm (Excel macro-enabled book)
- xltx (Excel template)
- xltm (Excel macro-enabled template)

## Installation

### Installing via NPM

excel-mcp-server is automatically installed by adding the following configuration to the MCP servers configuration.

For Windows:

```json
{
    "mcpServers": {
        "excel": {
            "command": "cmd",
            "args": ["/c", "npx", "--yes", "ms-excel-mcp-server"],
            "env": {
                "EXCEL_MCP_PAGING_CELLS_LIMIT": "4000"
            }
        }
    }
}
```

For other platforms:

```json
{
    "mcpServers": {
        "excel": {
            "command": "npx",
            "args": ["--yes", "ms-excel-mcp-server"],
            "env": {
                "EXCEL_MCP_PAGING_CELLS_LIMIT": "4000"
            }
        }
    }
}
```

### Installing via Smithery

To install Excel MCP Server for Claude Desktop automatically via [Smithery](https://smithery.ai/server/ms-excel-mcp-server):

```bash
npx -y @smithery/cli install @vKenjo/ms-excel-mcp-server --client claude
```

<h2 id="tools">Tools</h2>

### `excel_describe_sheets`

List all sheet information of specified Excel file.

**Arguments:**

- `fileAbsolutePath`
  - Absolute path to the Excel file

### `excel_read_sheet`

Read values from Excel sheet with pagination.

**Arguments:**

- `fileAbsolutePath`
  - Absolute path to the Excel file
- `sheetName`
  - Sheet name in the Excel file
- `range`
  - Range of cells to read in the Excel sheet (e.g., "A1:C10"). [default: first paging range]
- `showFormula`
  - Show formula instead of value [default: false]
- `showStyle`
  - Show style information for cells [default: false]

### `excel_screen_capture`

**[Windows only]** Take a screenshot of the Excel sheet with pagination.

**Arguments:**

- `fileAbsolutePath`
  - Absolute path to the Excel file
- `sheetName`
  - Sheet name in the Excel file
- `range`
  - Range of cells to read in the Excel sheet (e.g., "A1:C10"). [default: first paging range]

### `excel_write_to_sheet`

Write values to the Excel sheet.

**Arguments:**

- `fileAbsolutePath`
  - Absolute path to the Excel file
- `sheetName`
  - Sheet name in the Excel file
- `newSheet`
  - Create a new sheet if true, otherwise write to the existing sheet
- `range`
  - Range of cells to read in the Excel sheet (e.g., "A1:C10").
- `values`
  - Values to write to the Excel sheet. If the value is a formula, it should start with "="

### `excel_create_table`

Create a table in the Excel sheet

**Arguments:**

- `fileAbsolutePath`
  - Absolute path to the Excel file
- `sheetName`
  - Sheet name where the table is created
- `range`
  - Range to be a table (e.g., "A1:C10")
- `tableName`
  - Table name to be created

### `excel_copy_sheet`

Copy existing sheet to a new sheet

**Arguments:**

- `fileAbsolutePath`
  - Absolute path to the Excel file
- `srcSheetName`
  - Source sheet name in the Excel file
- `dstSheetName`
  - Sheet name to be copied

### `excel_add_data_validation`

Add data validation to Excel cells including dropdown lists and input validation.

**Arguments:**

- `fileAbsolutePath`
  - Absolute path to the Excel file
- `sheetName`
  - Sheet name in the Excel file
- `cellRange`
  - Range of cells to apply data validation (e.g., "A1:A10")
- `validationType`
  - Type of validation: list, whole, decimal, date, time, textLength, custom
- `options`
  - Data validation options (dropdownList, formulas, error messages, etc.)

### `excel_add_conditional_formatting`

Add conditional formatting to Excel cells with highlighting, color scales, and data bars.

**Arguments:**

- `fileAbsolutePath`
  - Absolute path to the Excel file
- `sheetName`
  - Sheet name in the Excel file  
- `cellRange`
  - Range of cells to apply conditional formatting (e.g., "A1:A10")
- `conditions`
  - Conditional formatting conditions (type, criteria, values, formatting)

### `excel_execute_vba` (Windows OLE only)

Execute VBA code on an Excel worksheet.

**Arguments:**

- `fileAbsolutePath`
  - Absolute path to the Excel file
- `sheetName`
  - Sheet name in the Excel file
- `vbaCode`
  - VBA code to execute

### `excel_add_vba_module` (Windows OLE only)

Add a VBA module to an Excel workbook.

**Arguments:**

- `fileAbsolutePath`
  - Absolute path to the Excel file
- `moduleName`
  - Name for the VBA module
- `vbaCode`
  - VBA code to add to the module

> **Note:** For detailed examples and usage instructions, see [docs/NEW_FEATURES.md](docs/NEW_FEATURES.md)

<h2 id="configuration">Configuration</h2>

You can change the MCP Server behaviors by the following environment variables:

### `EXCEL_MCP_PAGING_CELLS_LIMIT`

The maximum number of cells to read in a single paging operation.  
[default: 4000]

## License

Copyright (c) 2025 Kazuki Negoro

excel-mcp-server is released under the [MIT License](LICENSE)
