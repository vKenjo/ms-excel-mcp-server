package tools

import (
	"context"
	"fmt"

	z "github.com/Oudwins/zog"
	"github.com/mark3labs/mcp-go/mcp"
	"github.com/mark3labs/mcp-go/server"
	"github.com/negokaz/excel-mcp-server/internal/excel"
	imcp "github.com/negokaz/excel-mcp-server/internal/mcp"
)

type ExcelAddConditionalFormattingArguments struct {
	FileAbsolutePath string `zog:"fileAbsolutePath"`
	SheetName        string `zog:"sheetName"`
	CellRange        string `zog:"cellRange"`
}

var excelAddConditionalFormattingArgumentsSchema = z.Struct(z.Schema{
	"fileAbsolutePath": z.String().Test(AbsolutePathTest()).Required(),
	"sheetName":        z.String().Required(),
	"cellRange":        z.String().Required(),
})

func AddExcelAddConditionalFormattingTool(server *server.MCPServer) {
	server.AddTool(mcp.NewTool("excel_add_conditional_formatting",
		mcp.WithDescription("Add conditional formatting to Excel cells"),
		mcp.WithString("fileAbsolutePath",
			mcp.Required(),
			mcp.Description("Absolute path to the Excel file"),
		),
		mcp.WithString("sheetName",
			mcp.Required(),
			mcp.Description("Sheet name in the Excel file"),
		),
		mcp.WithString("cellRange",
			mcp.Required(),
			mcp.Description("Range of cells to apply conditional formatting (e.g., \"A1:A10\")"),
		),
		mcp.WithObject("conditions",
			mcp.Description("Conditional formatting conditions including type, criteria, values, and formatting"),
		),
	), handleAddConditionalFormatting)
}

func handleAddConditionalFormatting(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	var args ExcelAddConditionalFormattingArguments
	issues := excelAddConditionalFormattingArgumentsSchema.Parse(request.Params.Arguments, &args)
	if len(issues) != 0 {
		return imcp.NewToolResultZogIssueMap(issues), nil
	}

	// Parse conditions manually from request
	var conditions *excel.ConditionalFormattingConditions
	if conditionsArg, ok := request.Params.Arguments["conditions"].(map[string]interface{}); ok {
		conditions = &excel.ConditionalFormattingConditions{}

		if condType, ok := conditionsArg["type"].(string); ok {
			conditions.Type = condType
		}

		if criteria, ok := conditionsArg["criteria"].(string); ok {
			conditions.Criteria = criteria
		}

		if value1, ok := conditionsArg["value1"].(string); ok {
			conditions.Value1 = value1
		}

		if value2, ok := conditionsArg["value2"].(string); ok {
			conditions.Value2 = value2
		}

		if formula, ok := conditionsArg["formula"].(string); ok {
			conditions.Formula = formula
		}

		// Parse format object
		if formatArg, ok := conditionsArg["format"].(map[string]interface{}); ok {
			conditions.Format = &excel.ConditionalFormattingStyle{}

			// Parse font
			if fontArg, ok := formatArg["font"].(map[string]interface{}); ok {
				conditions.Format.Font = &excel.FontStyle{}
				if bold, ok := fontArg["bold"].(bool); ok {
					conditions.Format.Font.Bold = bold
				}
				if italic, ok := fontArg["italic"].(bool); ok {
					conditions.Format.Font.Italic = italic
				}
				if color, ok := fontArg["color"].(string); ok {
					conditions.Format.Font.Color = color
				}
				if size, ok := fontArg["size"].(float64); ok {
					conditions.Format.Font.Size = int(size)
				}
			}

			// Parse fill
			if fillArg, ok := formatArg["fill"].(map[string]interface{}); ok {
				conditions.Format.Fill = &excel.FillStyle{}
				if fillType, ok := fillArg["type"].(string); ok {
					conditions.Format.Fill.Type = fillType
				}
				if colorArg, ok := fillArg["color"].([]interface{}); ok {
					conditions.Format.Fill.Color = make([]string, len(colorArg))
					for i, c := range colorArg {
						if colorStr, ok := c.(string); ok {
							conditions.Format.Fill.Color[i] = colorStr
						}
					}
				}
			}
		}

		// Parse color scale options
		if colorScaleArg, ok := conditionsArg["colorScale"].(map[string]interface{}); ok {
			conditions.ColorScale = &excel.ColorScaleOptions{}
			if minType, ok := colorScaleArg["minType"].(string); ok {
				conditions.ColorScale.MinType = minType
			}
			if minValue, ok := colorScaleArg["minValue"].(string); ok {
				conditions.ColorScale.MinValue = minValue
			}
			if minColor, ok := colorScaleArg["minColor"].(string); ok {
				conditions.ColorScale.MinColor = minColor
			}
			if maxType, ok := colorScaleArg["maxType"].(string); ok {
				conditions.ColorScale.MaxType = maxType
			}
			if maxValue, ok := colorScaleArg["maxValue"].(string); ok {
				conditions.ColorScale.MaxValue = maxValue
			}
			if maxColor, ok := colorScaleArg["maxColor"].(string); ok {
				conditions.ColorScale.MaxColor = maxColor
			}
		}

		// Parse data bar options
		if dataBarArg, ok := conditionsArg["dataBar"].(map[string]interface{}); ok {
			conditions.DataBar = &excel.DataBarOptions{}
			if minType, ok := dataBarArg["minType"].(string); ok {
				conditions.DataBar.MinType = minType
			}
			if minValue, ok := dataBarArg["minValue"].(string); ok {
				conditions.DataBar.MinValue = minValue
			}
			if maxType, ok := dataBarArg["maxType"].(string); ok {
				conditions.DataBar.MaxType = maxType
			}
			if maxValue, ok := dataBarArg["maxValue"].(string); ok {
				conditions.DataBar.MaxValue = maxValue
			}
			if color, ok := dataBarArg["color"].(string); ok {
				conditions.DataBar.Color = color
			}
			if showValue, ok := dataBarArg["showValue"].(bool); ok {
				conditions.DataBar.ShowValue = showValue
			}
		}
	}

	workbook, releaseWorkbook, err := excel.OpenFile(args.FileAbsolutePath)
	if err != nil {
		return nil, err
	}
	defer releaseWorkbook()

	worksheet, err := workbook.FindSheet(args.SheetName)
	if err != nil {
		return imcp.NewToolResultInvalidArgumentError(fmt.Sprintf("sheet '%s' not found: %v", args.SheetName, err)), nil
	}
	defer worksheet.Release()

	err = worksheet.AddConditionalFormatting(args.CellRange, conditions)
	if err != nil {
		return nil, err
	}

	err = workbook.Save()
	if err != nil {
		return nil, err
	}

	return &mcp.CallToolResult{
		Content: []mcp.Content{
			mcp.NewTextContent(fmt.Sprintf("Conditional formatting successfully added to range %s in sheet '%s'", args.CellRange, args.SheetName)),
		},
	}, nil
}
