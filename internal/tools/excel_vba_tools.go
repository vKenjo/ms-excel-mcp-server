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

type ExcelExecuteVBAArguments struct {
	FileAbsolutePath string `zog:"fileAbsolutePath"`
	SheetName        string `zog:"sheetName"`
	VBACode          string `zog:"vbaCode"`
}

var excelExecuteVBAArgumentsSchema = z.Struct(z.Schema{
	"fileAbsolutePath": z.String().Test(AbsolutePathTest()).Required(),
	"sheetName":        z.String().Required(),
	"vbaCode":          z.String().Required(),
})

func AddExcelExecuteVBATool(server *server.MCPServer) {
	server.AddTool(mcp.NewTool("excel_execute_vba",
		mcp.WithDescription("Execute VBA code on an Excel worksheet"),
		mcp.WithString("fileAbsolutePath",
			mcp.Required(),
			mcp.Description("Absolute path to the Excel file"),
		),
		mcp.WithString("sheetName",
			mcp.Required(),
			mcp.Description("Sheet name in the Excel file"),
		),
		mcp.WithString("vbaCode",
			mcp.Required(),
			mcp.Description("VBA code to execute on the worksheet"),
		),
	), handleExecuteVBA)
}

func handleExecuteVBA(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	var args ExcelExecuteVBAArguments
	issues := excelExecuteVBAArgumentsSchema.Parse(request.Params.Arguments, &args)
	if len(issues) != 0 {
		return imcp.NewToolResultZogIssueMap(issues), nil
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

	err = worksheet.ExecuteVBA(args.VBACode)
	if err != nil {
		return nil, err
	}

	err = workbook.Save()
	if err != nil {
		return nil, err
	}

	return &mcp.CallToolResult{
		Content: []mcp.Content{
			mcp.NewTextContent(fmt.Sprintf("VBA code successfully executed on sheet '%s'", args.SheetName)),
		},
	}, nil
}

type ExcelAddVBAModuleArguments struct {
	FileAbsolutePath string `zog:"fileAbsolutePath"`
	ModuleName       string `zog:"moduleName"`
	VBACode          string `zog:"vbaCode"`
}

var excelAddVBAModuleArgumentsSchema = z.Struct(z.Schema{
	"fileAbsolutePath": z.String().Test(AbsolutePathTest()).Required(),
	"moduleName":       z.String().Required(),
	"vbaCode":          z.String().Required(),
})

func AddExcelAddVBAModuleTool(server *server.MCPServer) {
	server.AddTool(mcp.NewTool("excel_add_vba_module",
		mcp.WithDescription("Add a VBA module to an Excel workbook"),
		mcp.WithString("fileAbsolutePath",
			mcp.Required(),
			mcp.Description("Absolute path to the Excel file"),
		),
		mcp.WithString("moduleName",
			mcp.Required(),
			mcp.Description("Name for the VBA module"),
		),
		mcp.WithString("vbaCode",
			mcp.Required(),
			mcp.Description("VBA code to add to the module"),
		),
	), handleAddVBAModule)
}

func handleAddVBAModule(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	var args ExcelAddVBAModuleArguments
	issues := excelAddVBAModuleArgumentsSchema.Parse(request.Params.Arguments, &args)
	if len(issues) != 0 {
		return imcp.NewToolResultZogIssueMap(issues), nil
	}

	workbook, releaseWorkbook, err := excel.OpenFile(args.FileAbsolutePath)
	if err != nil {
		return nil, err
	}
	defer releaseWorkbook()

	// For VBA module, we need to get any worksheet to access the workbook
	sheets, err := workbook.GetSheets()
	if err != nil || len(sheets) == 0 {
		return imcp.NewToolResultInvalidArgumentError("no sheets found in workbook"), nil
	}

	firstSheetName, err := sheets[0].Name()
	if err != nil {
		return imcp.NewToolResultInvalidArgumentError("could not get sheet name"), nil
	}
	sheets[0].Release() // Release the sheet we got the name from

	worksheet, err := workbook.FindSheet(firstSheetName)
	if err != nil {
		return imcp.NewToolResultInvalidArgumentError("could not access workbook for VBA module"), nil
	}
	defer worksheet.Release()

	err = worksheet.AddVBAModule(args.ModuleName, args.VBACode)
	if err != nil {
		return nil, err
	}

	err = workbook.Save()
	if err != nil {
		return nil, err
	}

	return &mcp.CallToolResult{
		Content: []mcp.Content{
			mcp.NewTextContent(fmt.Sprintf("VBA module '%s' successfully added to workbook", args.ModuleName)),
		},
	}, nil
}
