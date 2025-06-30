# Installation Guide for ms-excel-mcp-server

## Quick Installation (Recommended)

The package is now published on npm as `ms-excel-mcp-server`. You can use it immediately with Claude Desktop.

### Method 1: Direct npm installation (Works immediately)

Add this configuration to your Claude Desktop MCP servers configuration:

**For Windows:**
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

**For macOS/Linux:**
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

### Method 2: Smithery installation (Coming soon)

Smithery typically takes a few hours to index new packages. Once indexed, you'll be able to install using:

```bash
npx -y @smithery/cli install ms-excel-mcp-server --client claude
```

*Note: If this command fails with "Server not found", it means Smithery hasn't indexed the package yet. Use Method 1 instead.*

## Verification

To verify the installation works, restart Claude Desktop and try asking:
"Can you help me read an Excel file?"

The Excel MCP Server should be loaded and ready to help with Excel operations.

## Features Available

- ✅ Read/Write Excel files (.xlsx, .xlsm, .xltx, .xltm)
- ✅ Create and manage worksheets
- ✅ Data validation with dropdown lists
- ✅ Conditional formatting (highlighting, color scales, data bars)
- ✅ VBA code execution (Windows only)
- ✅ Screenshot capture (Windows only)
- ✅ Table creation and management

## Troubleshooting

If you encounter issues:

1. **Package not found**: Make sure you're using `ms-excel-mcp-server` (not `@vKenjo/excel-mcp-server`)
2. **Command fails**: Try clearing npm cache: `npm cache clean --force`
3. **Permission issues**: Make sure Claude Desktop has permission to execute npm commands
4. **Node.js version**: Requires Node.js 20.x or later

## Configuration Options

You can customize the behavior using environment variables:

- `EXCEL_MCP_PAGING_CELLS_LIMIT`: Maximum cells to read per page (default: 4000)

Example with custom configuration:
```json
{
    "mcpServers": {
        "excel": {
            "command": "npx",
            "args": ["--yes", "ms-excel-mcp-server"],
            "env": {
                "EXCEL_MCP_PAGING_CELLS_LIMIT": "2000"
            }
        }
    }
}
```
