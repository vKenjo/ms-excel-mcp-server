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

type ExcelAddDataValidationArguments struct {
	FileAbsolutePath string `zog:"fileAbsolutePath"`
	SheetName        string `zog:"sheetName"`
	CellRange        string `zog:"cellRange"`
	ValidationType   string `zog:"validationType"`
}

var excelAddDataValidationArgumentsSchema = z.Struct(z.Schema{
	"fileAbsolutePath": z.String().Test(AbsolutePathTest()).Required(),
	"sheetName":        z.String().Required(),
	"cellRange":        z.String().Required(),
	"validationType":   z.String().Required(),
})

func AddExcelAddDataValidationTool(server *server.MCPServer) {
	server.AddTool(mcp.NewTool("excel_add_data_validation",
		mcp.WithDescription("Add data validation (including dropdown lists) to Excel cells"),
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
			mcp.Description("Range of cells to apply data validation (e.g., \"A1:A10\")"),
		),
		mcp.WithString("validationType",
			mcp.Required(),
			mcp.Description("Type of validation: list, whole, decimal, date, time, textLength, custom"),
		),
		mcp.WithObject("options",
			mcp.Description("Data validation options including dropdownList, formulas, error messages, etc."),
		),
	), handleAddDataValidation)
}

func handleAddDataValidation(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	var args ExcelAddDataValidationArguments
	issues := excelAddDataValidationArgumentsSchema.Parse(request.Params.Arguments, &args)
	if len(issues) != 0 {
		return imcp.NewToolResultZogIssueMap(issues), nil
	}

	// Parse options manually from request
	var options *excel.DataValidationOptions
	if optionsArg, ok := request.Params.Arguments["options"].(map[string]interface{}); ok {
		options = &excel.DataValidationOptions{}

		if dropdownList, ok := optionsArg["dropdownList"].([]interface{}); ok {
			options.DropdownList = make([]string, len(dropdownList))
			for i, item := range dropdownList {
				if str, ok := item.(string); ok {
					options.DropdownList[i] = str
				}
			}
		}

		if formula1, ok := optionsArg["formula1"].(string); ok {
			options.Formula1 = formula1
		}

		if formula2, ok := optionsArg["formula2"].(string); ok {
			options.Formula2 = formula2
		}

		if operator, ok := optionsArg["operator"].(string); ok {
			options.Operator = operator
		}

		if showErrorMessage, ok := optionsArg["showErrorMessage"].(bool); ok {
			options.ShowErrorMessage = showErrorMessage
		}

		if errorTitle, ok := optionsArg["errorTitle"].(string); ok {
			options.ErrorTitle = errorTitle
		}

		if errorMessage, ok := optionsArg["errorMessage"].(string); ok {
			options.ErrorMessage = errorMessage
		}

		if showInputMessage, ok := optionsArg["showInputMessage"].(bool); ok {
			options.ShowInputMessage = showInputMessage
		}

		if inputTitle, ok := optionsArg["inputTitle"].(string); ok {
			options.InputTitle = inputTitle
		}

		if inputMessage, ok := optionsArg["inputMessage"].(string); ok {
			options.InputMessage = inputMessage
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

	// Convert string validation type to enum
	var validationType excel.DataValidationType
	switch args.ValidationType {
	case "list":
		validationType = excel.DataValidationList
	case "whole":
		validationType = excel.DataValidationWhole
	case "decimal":
		validationType = excel.DataValidationDecimal
	case "date":
		validationType = excel.DataValidationDate
	case "time":
		validationType = excel.DataValidationTime
	case "textLength":
		validationType = excel.DataValidationTextLength
	case "custom":
		validationType = excel.DataValidationCustom
	default:
		return imcp.NewToolResultInvalidArgumentError(fmt.Sprintf("invalid validation type: %s", args.ValidationType)), nil
	}

	err = worksheet.AddDataValidation(args.CellRange, validationType, options)
	if err != nil {
		return nil, err
	}

	err = workbook.Save()
	if err != nil {
		return nil, err
	}

	return &mcp.CallToolResult{
		Content: []mcp.Content{
			mcp.NewTextContent(fmt.Sprintf("Data validation successfully added to range %s in sheet '%s'", args.CellRange, args.SheetName)),
		},
	}, nil
}
