package excel

import (
	"bufio"
	"bytes"
	"encoding/base64"
	"fmt"
	"io"
	"path/filepath"
	"strings"

	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
	"github.com/skanehira/clipboard-image"
)

type OleExcel struct {
	workbook *ole.IDispatch
}

type OleWorksheet struct {
	workbook  *ole.IDispatch
	worksheet *ole.IDispatch
}

func NewExcelOle(absolutePath string) (*OleExcel, func(), error) {
	ole.CoInitializeEx(0, ole.COINIT_MULTITHREADED)

	unknown, err := oleutil.GetActiveObject("Excel.Application")
	if err != nil {
		ole.CoUninitialize()
		return nil, func() {}, err
	}
	excel, err := unknown.QueryInterface(ole.IID_IDispatch)
	if err != nil {
		unknown.Release()
		ole.CoUninitialize()
		return nil, func() {}, err
	}
	oleutil.MustPutProperty(excel, "ScreenUpdating", false)
	oleutil.MustPutProperty(excel, "EnableEvents", false)
	workbooks := oleutil.MustGetProperty(excel, "Workbooks").ToIDispatch()
	c := oleutil.MustGetProperty(workbooks, "Count").Val
	for i := 1; i <= int(c); i++ {
		workbook := oleutil.MustGetProperty(workbooks, "Item", i).ToIDispatch()
		fullName := oleutil.MustGetProperty(workbook, "FullName").ToString()
		name := oleutil.MustGetProperty(workbook, "Name").ToString()
		if strings.HasPrefix(fullName, "https:") && name == filepath.Base(absolutePath) {
			// If a workbook is opened through a WOPI URL, its absolute file path cannot be retrieved.
			// If the absolutePath is not writable, it assumes that the workbook has opened by WOPI.
			if FileIsNotWritable(absolutePath) {
				return &OleExcel{workbook: workbook}, func() {
					oleutil.MustPutProperty(excel, "EnableEvents", true)
					oleutil.MustPutProperty(excel, "ScreenUpdating", true)
					workbook.Release()
					workbooks.Release()
					excel.Release()
					ole.CoUninitialize()
				}, nil
			} else {
				// This workbook might not be specified with the absolutePath
			}
		} else if normalizePath(fullName) == normalizePath(absolutePath) {
			return &OleExcel{workbook: workbook}, func() {
				oleutil.MustPutProperty(excel, "EnableEvents", true)
				oleutil.MustPutProperty(excel, "ScreenUpdating", true)
				workbook.Release()
				workbooks.Release()
				excel.Release()
				ole.CoUninitialize()
			}, nil
		}
	}
	return nil, func() {}, fmt.Errorf("workbook not found: %s", absolutePath)
}

func NewExcelOleWithNewObject(absolutePath string) (*OleExcel, func(), error) {
	ole.CoInitializeEx(0, ole.COINIT_MULTITHREADED)

	unknown, err := oleutil.CreateObject("Excel.Application")
	if err != nil {
		ole.CoUninitialize()
		return nil, func() {}, err
	}
	excel, err := unknown.QueryInterface(ole.IID_IDispatch)
	if err != nil {
		unknown.Release()
		ole.CoUninitialize()
		return nil, func() {}, err
	}
	workbooks := oleutil.MustGetProperty(excel, "Workbooks").ToIDispatch()
	workbook, err := oleutil.CallMethod(workbooks, "Open", absolutePath)
	if err != nil {
		workbooks.Release()
		excel.Release()
		ole.CoUninitialize()
		return nil, func() {}, err
	}
	w := workbook.ToIDispatch()
	return &OleExcel{workbook: w}, func() {
		w.Release()
		workbooks.Release()
		excel.Release()
		oleutil.CallMethod(excel, "Close")
		ole.CoUninitialize()
	}, nil
}

func (o *OleExcel) GetBackendName() string {
	return "ole"
}

func (o *OleExcel) GetSheets() ([]Worksheet, error) {
	worksheets := oleutil.MustGetProperty(o.workbook, "Worksheets").ToIDispatch()
	defer worksheets.Release()

	count := int(oleutil.MustGetProperty(worksheets, "Count").Val)
	worksheetList := make([]Worksheet, count)

	for i := 1; i <= count; i++ {
		worksheet := oleutil.MustGetProperty(worksheets, "Item", i).ToIDispatch()
		worksheetList[i-1] = &OleWorksheet{
			workbook:  o.workbook,
			worksheet: worksheet,
		}
	}
	return worksheetList, nil
}

func (o *OleExcel) FindSheet(sheetName string) (Worksheet, error) {
	worksheets := oleutil.MustGetProperty(o.workbook, "Worksheets").ToIDispatch()
	defer worksheets.Release()

	count := int(oleutil.MustGetProperty(worksheets, "Count").Val)

	for i := 1; i <= count; i++ {
		worksheet := oleutil.MustGetProperty(worksheets, "Item", i).ToIDispatch()
		name := oleutil.MustGetProperty(worksheet, "Name").ToString()

		if name == sheetName {
			return &OleWorksheet{
				workbook:  o.workbook,
				worksheet: worksheet,
			}, nil
		}
	}

	return nil, fmt.Errorf("sheet not found: %s", sheetName)
}

func (o *OleExcel) CreateNewSheet(sheetName string) error {
	activeWorksheet := oleutil.MustGetProperty(o.workbook, "ActiveSheet").ToIDispatch()
	defer activeWorksheet.Release()
	activeWorksheetIndex := oleutil.MustGetProperty(activeWorksheet, "Index").Val
	worksheets := oleutil.MustGetProperty(o.workbook, "Worksheets").ToIDispatch()
	defer worksheets.Release()

	_, err := oleutil.CallMethod(worksheets, "Add", nil, activeWorksheet)
	if err != nil {
		return err
	}

	worksheet := oleutil.MustGetProperty(worksheets, "Item", activeWorksheetIndex+1).ToIDispatch()
	defer worksheet.Release()

	_, err = oleutil.PutProperty(worksheet, "Name", sheetName)
	if err != nil {
		return err
	}

	return nil
}

func (o *OleExcel) CopySheet(srcSheetName string, dstSheetName string) error {
	worksheets := oleutil.MustGetProperty(o.workbook, "Worksheets").ToIDispatch()
	defer worksheets.Release()

	srcSheetVariant, err := oleutil.GetProperty(worksheets, "Item", srcSheetName)
	if err != nil {
		return fmt.Errorf("faild to get sheet: %w", err)
	}
	srcSheet := srcSheetVariant.ToIDispatch()
	defer srcSheet.Release()
	srcSheetIndex := oleutil.MustGetProperty(srcSheet, "Index").Val

	_, err = oleutil.CallMethod(srcSheet, "Copy", nil, srcSheet)
	if err != nil {
		return err
	}

	dstSheetVariant, err := oleutil.GetProperty(worksheets, "Item", srcSheetIndex+1)
	if err != nil {
		return fmt.Errorf("failed to get copied sheet: %w", err)
	}
	dstSheet := dstSheetVariant.ToIDispatch()
	defer dstSheet.Release()

	_, err = oleutil.PutProperty(dstSheet, "Name", dstSheetName)
	if err != nil {
		return err
	}

	return nil
}

func (o *OleExcel) Save() error {
	_, err := oleutil.CallMethod(o.workbook, "Save")
	if err != nil {
		return err
	}
	return nil
}

func (o *OleWorksheet) Release() {
	o.worksheet.Release()
}

func (o *OleWorksheet) Name() (string, error) {
	name := oleutil.MustGetProperty(o.worksheet, "Name").ToString()
	return name, nil
}

func (o *OleWorksheet) GetTables() ([]Table, error) {
	tables := oleutil.MustGetProperty(o.worksheet, "ListObjects").ToIDispatch()
	defer tables.Release()
	count := int(oleutil.MustGetProperty(tables, "Count").Val)
	tableList := make([]Table, count)
	for i := 1; i <= count; i++ {
		table := oleutil.MustGetProperty(tables, "Item", i).ToIDispatch()
		defer table.Release()
		name := oleutil.MustGetProperty(table, "Name").ToString()
		defer table.Release()
		tableRange := oleutil.MustGetProperty(table, "Range").ToIDispatch()
		defer tableRange.Release()
		tableList[i-1] = Table{
			Name:  name,
			Range: NormalizeRange(oleutil.MustGetProperty(tableRange, "Address").ToString()),
		}
	}
	return tableList, nil
}

func (o *OleWorksheet) GetPivotTables() ([]PivotTable, error) {
	pivotTables := oleutil.MustGetProperty(o.worksheet, "PivotTables").ToIDispatch()
	defer pivotTables.Release()
	count := int(oleutil.MustGetProperty(pivotTables, "Count").Val)
	pivotTableList := make([]PivotTable, count)
	for i := 1; i <= count; i++ {
		pivotTable := oleutil.MustGetProperty(pivotTables, "Item", i).ToIDispatch()
		defer pivotTable.Release()
		name := oleutil.MustGetProperty(pivotTable, "Name").ToString()
		pivotTableRange := oleutil.MustGetProperty(pivotTable, "TableRange1").ToIDispatch()
		defer pivotTableRange.Release()
		pivotTableList[i-1] = PivotTable{
			Name:  name,
			Range: NormalizeRange(oleutil.MustGetProperty(pivotTableRange, "Address").ToString()),
		}
	}
	return pivotTableList, nil
}

func (o *OleWorksheet) SetValue(cell string, value any) error {
	range_ := oleutil.MustGetProperty(o.worksheet, "Range", cell).ToIDispatch()
	defer range_.Release()
	_, err := oleutil.PutProperty(range_, "Value", value)
	return err
}

func (o *OleWorksheet) SetFormula(cell string, formula string) error {
	range_ := oleutil.MustGetProperty(o.worksheet, "Range", cell).ToIDispatch()
	defer range_.Release()
	_, err := oleutil.PutProperty(range_, "Formula", formula)
	return err
}

func (o *OleWorksheet) GetValue(cell string) (string, error) {
	range_ := oleutil.MustGetProperty(o.worksheet, "Range", cell).ToIDispatch()
	defer range_.Release()
	value := oleutil.MustGetProperty(range_, "Text").Value()
	switch v := value.(type) {
	case string:
		return v, nil
	case nil:
		return "", nil
	default: // Handle other types as needed
		return "", fmt.Errorf("unsupported type: %T", v)
	}
}

func (o *OleWorksheet) GetFormula(cell string) (string, error) {
	range_ := oleutil.MustGetProperty(o.worksheet, "Range", cell).ToIDispatch()
	defer range_.Release()
	formula := oleutil.MustGetProperty(range_, "Formula").ToString()
	return formula, nil
}

func (o *OleWorksheet) GetDimention() (string, error) {
	range_ := oleutil.MustGetProperty(o.worksheet, "UsedRange").ToIDispatch()
	defer range_.Release()
	dimension := oleutil.MustGetProperty(range_, "Address").ToString()
	return NormalizeRange(dimension), nil
}

func (o *OleWorksheet) GetPagingStrategy(pageSize int) (PagingStrategy, error) {
	return NewOlePagingStrategy(1000, o)
}

func (o *OleWorksheet) PrintArea() (string, error) {
	v, err := oleutil.GetProperty(o.worksheet, "PageSetup")
	if err != nil {
		return "", err
	}
	pageSetup := v.ToIDispatch()
	defer pageSetup.Release()

	printArea := oleutil.MustGetProperty(pageSetup, "PrintArea").ToString()
	return printArea, nil
}

func (o *OleWorksheet) HPageBreaks() ([]int, error) {
	v, err := oleutil.GetProperty(o.worksheet, "HPageBreaks")
	if err != nil {
		return nil, err
	}
	hPageBreaks := v.ToIDispatch()
	defer hPageBreaks.Release()

	count := int(oleutil.MustGetProperty(hPageBreaks, "Count").Val)
	pageBreaks := make([]int, count)
	for i := 1; i <= count; i++ {
		pageBreak := oleutil.MustGetProperty(hPageBreaks, "Item", i).ToIDispatch()
		defer pageBreak.Release()
		location := oleutil.MustGetProperty(pageBreak, "Location").ToIDispatch()
		defer location.Release()
		row := oleutil.MustGetProperty(location, "Row").Val
		pageBreaks[i-1] = int(row)
	}
	return pageBreaks, nil
}

func (o *OleWorksheet) CapturePicture(captureRange string) (string, error) {
	r := oleutil.MustGetProperty(o.worksheet, "Range", captureRange).ToIDispatch()
	defer r.Release()
	_, err := oleutil.CallMethod(
		r,
		"CopyPicture",
		int(1), // xlScreen (https://learn.microsoft.com/ja-jp/dotnet/api/microsoft.office.interop.excel.xlpictureappearance?view=excel-pia)
		int(2), // xlBitmap (https://learn.microsoft.com/ja-jp/dotnet/api/microsoft.office.interop.excel.xlcopypictureformat?view=excel-pia)
	)
	if err != nil {
		return "", err
	}
	// Read the image from the clipboard
	buf := new(bytes.Buffer)
	bufWriter := bufio.NewWriter(buf)
	clipboardReader, err := clipboard.ReadFromClipboard()
	if err != nil {
		return "", fmt.Errorf("failed to read from clipboard: %w", err)
	}
	if _, err := io.Copy(bufWriter, clipboardReader); err != nil {
		return "", fmt.Errorf("failed to copy clipboard data: %w", err)
	}
	if err := bufWriter.Flush(); err != nil {
		return "", fmt.Errorf("failed to flush buffer: %w", err)
	}
	return base64.StdEncoding.EncodeToString(buf.Bytes()), nil
}

func (o *OleWorksheet) AddTable(tableRange string, tableName string) error {
	tables := oleutil.MustGetProperty(o.worksheet, "ListObjects").ToIDispatch()
	defer tables.Release()

	// https://learn.microsoft.com/ja-jp/office/vba/api/excel.listobjects.add
	tableVar, err := oleutil.CallMethod(
		tables,
		"Add",
		int(1), // xlSrcRange (https://learn.microsoft.com/ja-jp/office/vba/api/excel.xllistobjectsourcetype)
		tableRange,
		nil,
		int(0), // xlYes (https://learn.microsoft.com/ja-jp/office/vba/api/excel.xlyesnoguess)
	)
	if err != nil {
		return err
	}
	table := tableVar.ToIDispatch()
	defer table.Release()
	_, err = oleutil.PutProperty(table, "Name", tableName)
	if err != nil {
		return err
	}
	return err
}

func (o *OleWorksheet) GetCellStyle(cell string) (*CellStyle, error) {
	rng := oleutil.MustGetProperty(o.worksheet, "Range", cell).ToIDispatch()
	defer rng.Release()

	style := &CellStyle{}

	// Get Font information
	font := oleutil.MustGetProperty(rng, "Font").ToIDispatch()
	defer font.Release()

	fontSize := int(oleutil.MustGetProperty(font, "Size").Value().(float64))
	fontBold := oleutil.MustGetProperty(font, "Bold").Value().(bool)
	fontItalic := oleutil.MustGetProperty(font, "Italic").Value().(bool)
	fontColor := oleutil.MustGetProperty(font, "Color").Value().(float64)

	style.Font = &FontStyle{
		Bold:   fontBold,
		Italic: fontItalic,
		Size:   fontSize,
		Color:  bgrToRgb(fontColor),
	}

	// Get Interior (fill) information
	interior := oleutil.MustGetProperty(rng, "Interior").ToIDispatch()
	defer interior.Release()

	interiorPattern := excelPatternToFillPattern(oleutil.MustGetProperty(interior, "Pattern").Value().(int32))

	if interiorPattern != FillPatternNone {
		interiorColor := oleutil.MustGetProperty(interior, "Color").Value().(float64)

		style.Fill = &FillStyle{
			Type:    "pattern",
			Pattern: interiorPattern,
			Color:   []string{bgrToRgb(interiorColor)},
		}
	}

	// Get Border information
	var borderStyles []BorderStyle

	// Get borders for each direction: Left(7), Top(8), Bottom(9), Right(10)
	borderPositions := []struct {
		index    int
		position string
	}{
		{7, "left"},
		{8, "top"},
		{9, "bottom"},
		{10, "right"},
	}

	for _, pos := range borderPositions {
		border := oleutil.MustGetProperty(rng, "Borders", pos.index).ToIDispatch()
		defer border.Release()

		borderLineStyle := excelBorderStyleToName(oleutil.MustGetProperty(border, "LineStyle").Value().(int32))

		if borderLineStyle != BorderStyleNone {
			borderColor := oleutil.MustGetProperty(border, "Color").Value().(float64)
			borderStyle := BorderStyle{
				Type:  pos.position,
				Style: borderLineStyle,
				Color: bgrToRgb(borderColor),
			}
			borderStyles = append(borderStyles, borderStyle)
		}
	}
	style.Border = borderStyles

	return style, nil
}

// bgrToRgb converts BGR color format to RGB hex string
func bgrToRgb(bgrColor float64) string {
	bgrColorInt := int32(bgrColor)
	// Extract RGB components from BGR format
	r := (bgrColorInt >> 0) & 0xFF
	g := (bgrColorInt >> 8) & 0xFF
	b := (bgrColorInt >> 16) & 0xFF
	return fmt.Sprintf("#%02X%02X%02X", r, g, b)
}

// excelBorderStyleToName converts Excel border style constant to BorderStyleName
func excelBorderStyleToName(excelStyle int32) BorderStyleName {
	switch excelStyle {
	case 1: // xlContinuous
		return BorderStyleContinuous
	case -4115: // xlDash
		return BorderStyleDash
	case -4118: // xlDot
		return BorderStyleDot
	case -4119: // xlDouble
		return BorderStyleDouble
	case 4: // xlDashDot
		return BorderStyleDashDot
	case 5: // xlDashDotDot
		return BorderStyleDashDotDot
	case 13: // xlSlantDashDot
		return BorderStyleSlantDashDot
	case -4142: // xlLineStyleNone
		return BorderStyleNone
	default:
		return BorderStyleNone
	}
}

// excelPatternToFillPattern converts Excel XlPattern constant to FillPatternName
func excelPatternToFillPattern(excelPattern int32) FillPatternName {
	switch excelPattern {
	case -4142: // xlPatternNone
		return FillPatternNone
	case 1: // xlPatternSolid
		return FillPatternSolid
	case -4125: // xlPatternGray75
		return FillPatternDarkGray
	case -4124: // xlPatternGray50
		return FillPatternMediumGray
	case -4126: // xlPatternGray25
		return FillPatternLightGray
	case -4121: // xlPatternGray16
		return FillPatternGray125
	case -4127: // xlPatternGray8
		return FillPatternGray0625
	case 9: // xlPatternHorizontal
		return FillPatternLightHorizontal
	case 12: // xlPatternVertical
		return FillPatternLightVertical
	case 10: // xlPatternDown
		return FillPatternLightDown
	case 11: // xlPatternUp
		return FillPatternLightUp
	case 16: // xlPatternGrid
		return FillPatternLightGrid
	case 17: // xlPatternCrissCross
		return FillPatternLightTrellis
	case 5: // xlPatternLightHorizontal
		return FillPatternLightHorizontal
	case 6: // xlPatternLightVertical
		return FillPatternLightVertical
	case 7: // xlPatternLightDown
		return FillPatternLightDown
	case 8: // xlPatternLightUp
		return FillPatternLightUp
	case 15: // xlPatternLightGrid
		return FillPatternLightGrid
	case 18: // xlPatternLightTrellis
		return FillPatternLightTrellis
	case 13: // xlPatternSemiGray75
		return FillPatternDarkHorizontal
	case 2: // xlPatternDarkHorizontal
		return FillPatternDarkHorizontal
	case 3: // xlPatternDarkVertical
		return FillPatternDarkVertical
	case 4: // xlPatternDarkDown
		return FillPatternDarkDown
	case 14: // xlPatternDarkUp
		return FillPatternDarkUp
	case -4162: // xlPatternDarkGrid
		return FillPatternDarkGrid
	case -4166: // xlPatternDarkTrellis
		return FillPatternDarkTrellis
	default:
		return FillPatternNone
	}
}

func normalizePath(path string) string {
	// Normalize the volume name to uppercase
	vol := filepath.VolumeName(path)
	if vol == "" {
		return path
	}
	rest := path[len(vol):]
	return filepath.Clean(strings.ToUpper(vol) + rest)
}

// AddDataValidation adds data validation to the specified range using OLE
func (o *OleWorksheet) AddDataValidation(cellRange string, validationType DataValidationType, options *DataValidationOptions) error {
	if options == nil {
		return fmt.Errorf("data validation options cannot be nil")
	}

	rng := oleutil.MustGetProperty(o.worksheet, "Range", cellRange).ToIDispatch()
	defer rng.Release()

	validation := oleutil.MustGetProperty(rng, "Validation").ToIDispatch()
	defer validation.Release()

	switch validationType {
	case DataValidationList:
		if len(options.DropdownList) > 0 {
			// Create dropdown list
			listString := strings.Join(options.DropdownList, ",")
			oleutil.MustCallMethod(validation, "Add", 3, listString, 7) // xlValidateList = 3, xlValidAlertStop = 1, xlBetween = 1
		}
	case DataValidationWhole:
		if options.Formula1 != "" {
			operator := getOleOperator(options.Operator)
			oleutil.MustCallMethod(validation, "Add", 1, operator, 7, options.Formula1, options.Formula2) // xlValidateWholeNumber = 1
		}
	case DataValidationDecimal:
		if options.Formula1 != "" {
			operator := getOleOperator(options.Operator)
			oleutil.MustCallMethod(validation, "Add", 2, operator, 7, options.Formula1, options.Formula2) // xlValidateDecimal = 2
		}
	case DataValidationDate:
		if options.Formula1 != "" {
			operator := getOleOperator(options.Operator)
			oleutil.MustCallMethod(validation, "Add", 4, operator, 7, options.Formula1, options.Formula2) // xlValidateDate = 4
		}
	case DataValidationTime:
		if options.Formula1 != "" {
			operator := getOleOperator(options.Operator)
			oleutil.MustCallMethod(validation, "Add", 5, operator, 7, options.Formula1, options.Formula2) // xlValidateTime = 5
		}
	case DataValidationTextLength:
		if options.Formula1 != "" {
			operator := getOleOperator(options.Operator)
			oleutil.MustCallMethod(validation, "Add", 6, operator, 7, options.Formula1, options.Formula2) // xlValidateTextLength = 6
		}
	case DataValidationCustom:
		if options.Formula1 != "" {
			oleutil.MustCallMethod(validation, "Add", 7, 1, 7, options.Formula1) // xlValidateCustom = 7
		}
	}

	// Set input and error messages if provided
	if options.ShowInputMessage && options.InputTitle != "" {
		oleutil.MustPutProperty(validation, "InputTitle", options.InputTitle)
		oleutil.MustPutProperty(validation, "InputMessage", options.InputMessage)
		oleutil.MustPutProperty(validation, "ShowInput", true)
	}

	if options.ShowErrorMessage && options.ErrorTitle != "" {
		oleutil.MustPutProperty(validation, "ErrorTitle", options.ErrorTitle)
		oleutil.MustPutProperty(validation, "ErrorMessage", options.ErrorMessage)
		oleutil.MustPutProperty(validation, "ShowError", true)
	}

	return nil
}

// getOleOperator converts string operator to OLE validation operator constant
func getOleOperator(operator string) int {
	switch operator {
	case "between":
		return 1 // xlBetween
	case "notBetween":
		return 2 // xlNotBetween
	case "equal":
		return 3 // xlEqual
	case "notEqual":
		return 4 // xlNotEqual
	case "greaterThan":
		return 5 // xlGreater
	case "lessThan":
		return 6 // xlLess
	case "greaterThanOrEqual":
		return 7 // xlGreaterEqual
	case "lessThanOrEqual":
		return 8 // xlLessEqual
	default:
		return 1 // xlBetween
	}
}

// AddConditionalFormatting adds conditional formatting to the specified range using OLE
func (o *OleWorksheet) AddConditionalFormatting(cellRange string, conditions *ConditionalFormattingConditions) error {
	if conditions == nil {
		return fmt.Errorf("conditional formatting conditions cannot be nil")
	}

	rng := oleutil.MustGetProperty(o.worksheet, "Range", cellRange).ToIDispatch()
	defer rng.Release()

	// Clear existing conditional formatting
	formatConditions := oleutil.MustGetProperty(rng, "FormatConditions").ToIDispatch()
	defer formatConditions.Release()
	oleutil.MustCallMethod(formatConditions, "Delete")

	switch conditions.Type {
	case "cellValue":
		operator := getOleConditionalOperator(conditions.Criteria)
		var condition *ole.IDispatch

		if conditions.Value2 != "" {
			condition = oleutil.MustCallMethod(formatConditions, "Add", 1, operator, conditions.Value1, conditions.Value2).ToIDispatch() // xlCellValue = 1
		} else {
			condition = oleutil.MustCallMethod(formatConditions, "Add", 1, operator, conditions.Value1).ToIDispatch()
		}
		defer condition.Release()

		// Apply formatting
		if conditions.Format != nil {
			applyOleFormatting(condition, conditions.Format)
		}

	case "expression":
		condition := oleutil.MustCallMethod(formatConditions, "Add", 2, conditions.Formula).ToIDispatch() // xlExpression = 2
		defer condition.Release()

		if conditions.Format != nil {
			applyOleFormatting(condition, conditions.Format)
		}

	case "colorScale":
		if conditions.ColorScale != nil {
			// Add color scale (Excel 2007+)
			colorScale := oleutil.MustCallMethod(formatConditions, "AddColorScale", 2).ToIDispatch() // 2 colors
			defer colorScale.Release()

			colorCriteria := oleutil.MustGetProperty(colorScale, "ColorScaleCriteria").ToIDispatch()
			defer colorCriteria.Release()

			// Set minimum
			minCriterion := oleutil.MustGetProperty(colorCriteria, "Item", 1).ToIDispatch()
			defer minCriterion.Release()
			oleutil.MustPutProperty(minCriterion, "Type", getOleColorScaleType(conditions.ColorScale.MinType))
			if conditions.ColorScale.MinValue != "" {
				oleutil.MustPutProperty(minCriterion, "Value", conditions.ColorScale.MinValue)
			}
			r, g, b := parseRGBColor(conditions.ColorScale.MinColor)
			app := oleutil.MustGetProperty(o.worksheet, "Application").ToIDispatch()
			defer app.Release()
			color := oleutil.MustGetProperty(app, "RGB", r, g, b).Value()
			oleutil.MustPutProperty(minCriterion, "FormatColor", color)

			// Set maximum
			maxCriterion := oleutil.MustGetProperty(colorCriteria, "Item", 2).ToIDispatch()
			defer maxCriterion.Release()
			oleutil.MustPutProperty(maxCriterion, "Type", getOleColorScaleType(conditions.ColorScale.MaxType))
			if conditions.ColorScale.MaxValue != "" {
				oleutil.MustPutProperty(maxCriterion, "Value", conditions.ColorScale.MaxValue)
			}
			r2, g2, b2 := parseRGBColor(conditions.ColorScale.MaxColor)
			color2 := oleutil.MustGetProperty(app, "RGB", r2, g2, b2).Value()
			oleutil.MustPutProperty(maxCriterion, "FormatColor", color2)
		}

	case "dataBar":
		if conditions.DataBar != nil {
			dataBar := oleutil.MustCallMethod(formatConditions, "AddDatabar").ToIDispatch()
			defer dataBar.Release()

			if conditions.DataBar.Color != "" {
				r, g, b := parseRGBColor(conditions.DataBar.Color)
				app := oleutil.MustGetProperty(o.worksheet, "Application").ToIDispatch()
				defer app.Release()
				color := oleutil.MustGetProperty(app, "RGB", r, g, b).Value()
				oleutil.MustPutProperty(dataBar, "BarColor", color)
			}
		}
	}

	return nil
}

// getOleConditionalOperator converts string criteria to OLE conditional operator constant
func getOleConditionalOperator(criteria string) int {
	switch criteria {
	case "greaterThan":
		return 5 // xlGreater
	case "lessThan":
		return 6 // xlLess
	case "between":
		return 1 // xlBetween
	case "equal":
		return 3 // xlEqual
	case "notEqual":
		return 4 // xlNotEqual
	case "greaterThanOrEqual":
		return 7 // xlGreaterEqual
	case "lessThanOrEqual":
		return 8 // xlLessEqual
	default:
		return 3 // xlEqual
	}
}

// getOleColorScaleType converts string type to OLE color scale type constant
func getOleColorScaleType(scaleType string) int {
	switch scaleType {
	case "num":
		return 0 // xlConditionValueNumber
	case "percent":
		return 1 // xlConditionValuePercent
	case "percentile":
		return 2 // xlConditionValuePercentile
	case "formula":
		return 3 // xlConditionValueFormula
	case "min":
		return 4 // xlConditionValueLowestValue
	case "max":
		return 5 // xlConditionValueHighestValue
	default:
		return 0 // xlConditionValueNumber
	}
}

// parseRGBColor parses hex color string to RGB values for OLE
func parseRGBColor(hexColor string) (int, int, int) {
	if len(hexColor) == 7 && hexColor[0] == '#' {
		r, g, b := 0, 0, 0
		fmt.Sscanf(hexColor[1:3], "%02x", &r)
		fmt.Sscanf(hexColor[3:5], "%02x", &g)
		fmt.Sscanf(hexColor[5:7], "%02x", &b)
		return r, g, b
	}
	return 0, 0, 0 // Default to black
}

// applyOleFormatting applies formatting to a conditional format condition
func applyOleFormatting(condition *ole.IDispatch, format *ConditionalFormattingStyle) {
	if format.Font != nil {
		font := oleutil.MustGetProperty(condition, "Font").ToIDispatch()
		defer font.Release()

		if format.Font.Bold {
			oleutil.MustPutProperty(font, "Bold", true)
		}
		if format.Font.Italic {
			oleutil.MustPutProperty(font, "Italic", true)
		}
		if format.Font.Color != "" {
			r, g, b := parseRGBColor(format.Font.Color)
			app := oleutil.MustGetProperty(condition, "Application").ToIDispatch()
			defer app.Release()
			color := oleutil.MustGetProperty(app, "RGB", r, g, b).Value()
			oleutil.MustPutProperty(font, "Color", color)
		}
		if format.Font.Size > 0 {
			oleutil.MustPutProperty(font, "Size", format.Font.Size)
		}
	}

	if format.Fill != nil && len(format.Fill.Color) > 0 {
		interior := oleutil.MustGetProperty(condition, "Interior").ToIDispatch()
		defer interior.Release()

		r, g, b := parseRGBColor(format.Fill.Color[0])
		app := oleutil.MustGetProperty(condition, "Application").ToIDispatch()
		defer app.Release()
		color := oleutil.MustGetProperty(app, "RGB", r, g, b).Value()
		oleutil.MustPutProperty(interior, "Color", color)
	}
}

// ExecuteVBA executes VBA code on the worksheet using OLE
func (o *OleWorksheet) ExecuteVBA(vbaCode string) error {
	app := oleutil.MustGetProperty(o.workbook, "Application").ToIDispatch()
	defer app.Release()

	// Execute VBA code using Application.Run or evaluate
	_, err := oleutil.CallMethod(app, "Evaluate", vbaCode)
	if err != nil {
		return fmt.Errorf("failed to execute VBA code: %w", err)
	}

	return nil
}

// AddVBAModule adds a VBA module to the workbook using OLE
func (o *OleWorksheet) AddVBAModule(moduleName, vbaCode string) error {
	vbProject := oleutil.MustGetProperty(o.workbook, "VBProject").ToIDispatch()
	defer vbProject.Release()

	vbComponents := oleutil.MustGetProperty(vbProject, "VBComponents").ToIDispatch()
	defer vbComponents.Release()

	// Add a new standard module (vbext_ct_StdModule = 1)
	newModule := oleutil.MustCallMethod(vbComponents, "Add", 1).ToIDispatch()
	defer newModule.Release()

	// Set the module name
	oleutil.MustPutProperty(newModule, "Name", moduleName)

	// Get the code module
	codeModule := oleutil.MustGetProperty(newModule, "CodeModule").ToIDispatch()
	defer codeModule.Release()

	// Add the VBA code
	lines := strings.Split(vbaCode, "\n")
	for i, line := range lines {
		oleutil.MustCallMethod(codeModule, "InsertLines", i+1, line)
	}

	return nil
}
