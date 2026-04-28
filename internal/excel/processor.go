package excel

import (
	"bytes"
	"errors"
	"fmt"
	"io"
	"math"
	"regexp"
	"sort"
	"strconv"
	"strings"
	"time"

	"github.com/shakinm/xlsReader/xls"
	"github.com/xuri/excelize/v2"
)

type SourceType string

const (
	SourceTypeOutbound SourceType = "outbound"
	SourceTypeWeidian  SourceType = "weidian"
)

var (
	ErrUnsupportedSourceType  = errors.New("unsupported source type")
	ErrMappingsNotImplemented = errors.New("excel field mappings are not implemented yet")
	ErrInvalidWorkbook        = errors.New("invalid workbook")
)

var productDetailRE = regexp.MustCompile(`产品名称[:：]\s*([^,;；]+)\s*[,，]\s*产品数量[:：]\s*([0-9]+(?:\.[0-9]+)?)`)

type Result struct {
	Workbook    *bytes.Buffer
	LogMarkdown string
}

type outboundRecord struct {
	Customer      string
	Applicant     string
	DocumentNo    string
	ContractNo    string
	Contact       string
	ApplyDate     string
	ProductCounts map[string]float64
	Products      []productLogEntry
	SourceRow     int
	SheetName     string
	OutputRow     int
}

type sourceRow struct {
	values []string
}

type headerIndex map[string]int

type productLogEntry struct {
	Raw         string
	SourceName  string
	Quantity    float64
	GoodsName   string
	OutputCol   int
	OutputCell  string
	Status      string
	Description string
}

type productMappingResult struct {
	Mapping map[string]string
	Rows    int
}

var outputProducts = []string{
	"用印工作台SS1(标准版)",
	"用印工作台SS1(Lite)",
	"智能终端S2Lite版(铜)",
	"智能终端S2Lite版(光敏)",
	"智能终端S2P(铜)",
	"智能终端S2P(光敏)",
	"章管家 MS1(铜)",
	"章管家 MS1(光敏)",
	"一体章S2AIO",
	"SC1A(主柜)",
	"SC1B(副柜)",
	"SC2M(主柜)",
	"SC2R(副柜)",
	"FC1A(主柜)",
	"FC1B(副柜)",
	"扫描仪",
	"多功能一体机",
	"工作站mini（主机）",
	"工作站mini（格口）",
	"工作站SC2C (小柜)",
}

var outputProductColumn = map[string]int{
	"用印工作台SS1(标准版)":   8,
	"用印工作台SS1(Lite)":  9,
	"智能终端S2Lite版(铜)":  10,
	"智能终端S2Lite版(光敏)": 11,
	"智能终端S2P(铜)":      12,
	"智能终端S2P(光敏)":     13,
	"章管家 MS1(铜)":      14,
	"章管家 MS1(光敏)":     15,
	"一体章S2AIO":        16,
	"SC1A(主柜)":        17,
	"SC1B(副柜)":        18,
	"SC2M(主柜)":        19,
	"SC2R(副柜)":        20,
	"FC1A(主柜)":        21,
	"FC1B(副柜)":        22,
	"扫描仪":             23,
	"多功能一体机":          24,
	"工作站mini（主机）":     25,
	"工作站mini（格口）":     26,
	"工作站SC2C (小柜)":    27,
}

func (s SourceType) Valid() bool {
	return s == SourceTypeOutbound || s == SourceTypeWeidian
}

func Transform(input io.Reader, sourceType SourceType, mappingInput io.Reader) (*Result, error) {
	if !sourceType.Valid() {
		return nil, fmt.Errorf("%w: %s", ErrUnsupportedSourceType, sourceType)
	}

	if sourceType != SourceTypeOutbound {
		return nil, ErrMappingsNotImplemented
	}

	inputBytes, err := io.ReadAll(input)
	if err != nil {
		return nil, fmt.Errorf("read workbook: %w", err)
	}

	productMapping, err := readProductMapping(mappingInput)
	if err != nil {
		return nil, err
	}

	records, err := parseOutboundWorkbook(inputBytes, productMapping)
	if err != nil {
		return nil, err
	}
	assignOutboundOutputPositions(records)

	workbook, err := buildOutboundWorkbook(records)
	if err != nil {
		return nil, err
	}

	return &Result{
		Workbook:    workbook,
		LogMarkdown: buildOutboundLog(records, productMapping.Rows),
	}, nil
}

func readProductMapping(input io.Reader) (*productMappingResult, error) {
	workbook, err := excelize.OpenReader(input)
	if err != nil {
		return nil, fmt.Errorf("open product mapping workbook: %w", err)
	}
	defer workbook.Close()

	sheets := workbook.GetSheetList()
	if len(sheets) == 0 {
		return nil, fmt.Errorf("%w: product mapping workbook has no sheets", ErrInvalidWorkbook)
	}

	rows, err := workbook.GetRows(sheets[0])
	if err != nil {
		return nil, fmt.Errorf("read product mapping rows: %w", err)
	}
	if len(rows) < 2 {
		return nil, fmt.Errorf("%w: product mapping workbook has no data", ErrInvalidWorkbook)
	}

	headers := indexHeaders(rows[0])
	sourceCol, ok := headers["出货产品名称"]
	if !ok {
		return nil, fmt.Errorf("%w: product mapping missing 出货产品名称", ErrInvalidWorkbook)
	}
	goodsCol, ok := headers["货品名称"]
	if !ok {
		return nil, fmt.Errorf("%w: product mapping missing 货品名称", ErrInvalidWorkbook)
	}

	mapping := make(map[string]string)
	for _, row := range rows[1:] {
		sourceName := valueAt(row, sourceCol)
		goodsName := valueAt(row, goodsCol)
		if sourceName == "" || goodsName == "" {
			continue
		}
		mapping[sourceName] = goodsName
	}
	if len(mapping) == 0 {
		return nil, fmt.Errorf("%w: product mapping has no usable rows", ErrInvalidWorkbook)
	}
	return &productMappingResult{Mapping: mapping, Rows: len(mapping)}, nil
}

func parseOutboundWorkbook(input []byte, productMapping *productMappingResult) ([]outboundRecord, error) {
	rows, err := readWorkbookRows(input)
	if err != nil {
		return nil, err
	}

	headerRowIndex := -1
	var headers headerIndex
	for i, row := range rows {
		current := indexHeaders(row.values)
		if _, ok := current["产品明细"]; ok {
			headerRowIndex = i
			headers = current
			break
		}
	}
	if headerRowIndex == -1 {
		return nil, fmt.Errorf("%w: outbound workbook missing 产品明细 header", ErrInvalidWorkbook)
	}

	requiredHeaders := []string{"客户名称", "申请人", "单据编号", "合同编号", "收货人", "收货电话", "申请时间", "产品明细"}
	for _, name := range requiredHeaders {
		if _, ok := headers[name]; !ok {
			return nil, fmt.Errorf("%w: outbound workbook missing %s header", ErrInvalidWorkbook, name)
		}
	}

	records := make([]outboundRecord, 0, len(rows)-headerRowIndex-1)
	for rowIndex, row := range rows[headerRowIndex+1:] {
		documentNo := row.get(headers["单据编号"])
		if documentNo == "" {
			continue
		}

		contact := joinNonEmpty(row.get(headers["收货人"]), cleanPhone(row.get(headers["收货电话"])))
		productCounts, productLogEntries := parseProductDetails(row.get(headers["产品明细"]), productMapping.Mapping)
		record := outboundRecord{
			Customer:      row.get(headers["客户名称"]),
			Applicant:     row.get(headers["申请人"]),
			DocumentNo:    documentNo,
			ContractNo:    row.get(headers["合同编号"]),
			Contact:       contact,
			ApplyDate:     row.get(headers["申请时间"]),
			ProductCounts: productCounts,
			Products:      productLogEntries,
			SourceRow:     headerRowIndex + rowIndex + 2,
		}
		records = append(records, record)
	}
	return records, nil
}

func assignOutboundOutputPositions(records []outboundRecord) {
	nextRowBySheet := make(map[string]int)
	for i := range records {
		sheetName := sheetNameFromDate(records[i].ApplyDate)
		nextRow := nextRowBySheet[sheetName]
		if nextRow == 0 {
			nextRow = 6
		}
		records[i].SheetName = sheetName
		records[i].OutputRow = nextRow
		for productIndex := range records[i].Products {
			outputCol := records[i].Products[productIndex].OutputCol
			if outputCol == 0 {
				continue
			}
			records[i].Products[productIndex].OutputCell = fmt.Sprintf("%s%d", columnName(outputCol), nextRow)
		}
		nextRowBySheet[sheetName] = nextRow + 1
	}
}

func readWorkbookRows(input []byte) ([]sourceRow, error) {
	if bytes.HasPrefix(input, []byte("PK")) {
		return readXLSXRows(input)
	}
	return readXLSRows(input)
}

func readXLSRows(input []byte) ([]sourceRow, error) {
	workbook, err := xls.OpenReader(bytes.NewReader(input))
	if err != nil {
		return nil, fmt.Errorf("open xls workbook: %w", err)
	}
	if workbook.GetNumberSheets() == 0 {
		return nil, fmt.Errorf("%w: workbook has no sheets", ErrInvalidWorkbook)
	}

	sheet, err := workbook.GetSheet(0)
	if err != nil {
		return nil, fmt.Errorf("read first xls sheet: %w", err)
	}

	rows := make([]sourceRow, 0, sheet.GetNumberRows())
	for rowIndex := 0; rowIndex < sheet.GetNumberRows(); rowIndex++ {
		row, err := sheet.GetRow(rowIndex)
		if err != nil || row == nil {
			rows = append(rows, sourceRow{})
			continue
		}
		values := make([]string, 256)
		lastValueIndex := -1
		for colIndex := 0; colIndex < len(values); colIndex++ {
			cell, err := row.GetCol(colIndex)
			if err != nil {
				continue
			}
			values[colIndex] = strings.TrimSpace(cell.GetString())
			if values[colIndex] != "" {
				lastValueIndex = colIndex
			}
		}
		if lastValueIndex >= 0 {
			values = values[:lastValueIndex+1]
		} else {
			values = nil
		}
		rows = append(rows, sourceRow{values: values})
	}
	return rows, nil
}

func readXLSXRows(input []byte) ([]sourceRow, error) {
	workbook, err := excelize.OpenReader(bytes.NewReader(input))
	if err != nil {
		return nil, fmt.Errorf("open xlsx workbook: %w", err)
	}
	defer workbook.Close()

	sheets := workbook.GetSheetList()
	if len(sheets) == 0 {
		return nil, fmt.Errorf("%w: workbook has no sheets", ErrInvalidWorkbook)
	}
	values, err := workbook.GetRows(sheets[0])
	if err != nil {
		return nil, fmt.Errorf("read xlsx rows: %w", err)
	}

	rows := make([]sourceRow, 0, len(values))
	for _, row := range values {
		cells := make([]string, len(row))
		for i, value := range row {
			cells[i] = strings.TrimSpace(value)
		}
		rows = append(rows, sourceRow{values: cells})
	}
	return rows, nil
}

func parseProductDetails(details string, productMapping map[string]string) (map[string]float64, []productLogEntry) {
	counts := make(map[string]float64)
	entries := make([]productLogEntry, 0)
	for _, part := range strings.FieldsFunc(details, func(r rune) bool { return r == ';' || r == '；' }) {
		part = strings.TrimSpace(part)
		if part == "" {
			continue
		}
		entry := productLogEntry{Raw: part}
		match := productDetailRE.FindStringSubmatch(part)
		if len(match) != 3 {
			entry.Status = "未解析"
			entry.Description = "产品明细片段不符合“产品名称:...,产品数量:...”格式，已忽略"
			entries = append(entries, entry)
			continue
		}
		sourceName := strings.TrimSpace(match[1])
		entry.SourceName = sourceName
		goodsName, ok := productMapping[sourceName]
		if !ok {
			entry.Status = "未映射"
			entry.Description = "硬件产品信息.xlsx 中未找到该出货产品名称，已忽略"
			quantity, _ := strconv.ParseFloat(match[2], 64)
			entry.Quantity = quantity
			entries = append(entries, entry)
			continue
		}
		quantity, err := strconv.ParseFloat(match[2], 64)
		if err != nil {
			entry.Status = "未解析"
			entry.Description = "产品数量不是有效数字，已忽略"
			entries = append(entries, entry)
			continue
		}
		entry.Quantity = quantity
		entry.GoodsName = goodsName
		if goodsName == "无" {
			entry.Status = "忽略"
			entry.Description = "硬件产品信息.xlsx 中货品名称为“无”，已忽略"
			entries = append(entries, entry)
			continue
		}
		outputCol, ok := outputProductColumn[goodsName]
		if !ok {
			entry.Status = "未写入"
			entry.Description = "映射后的货品名称不在处理后模板输出列中，已忽略"
			entries = append(entries, entry)
			continue
		}
		entry.OutputCol = outputCol
		entry.Status = "写入"
		entry.Description = "已按货品名称汇总到输出列"
		counts[goodsName] += quantity
		entries = append(entries, entry)
	}
	return counts, entries
}

func buildOutboundWorkbook(records []outboundRecord) (*bytes.Buffer, error) {
	workbook := excelize.NewFile()
	defer workbook.Close()

	recordsBySheet, sheetOrder := groupRecordsByMonth(records)
	defaultSheet := workbook.GetSheetName(workbook.GetActiveSheetIndex())
	for i, sheetName := range sheetOrder {
		if i == 0 {
			if err := workbook.SetSheetName(defaultSheet, sheetName); err != nil {
				return nil, fmt.Errorf("rename sheet: %w", err)
			}
		} else if _, err := workbook.NewSheet(sheetName); err != nil {
			return nil, fmt.Errorf("create sheet %s: %w", sheetName, err)
		}
		if err := writeOutboundSheet(workbook, sheetName, recordsBySheet[sheetName]); err != nil {
			return nil, err
		}
	}
	if len(sheetOrder) > 0 {
		index, err := workbook.GetSheetIndex(sheetOrder[0])
		if err != nil {
			return nil, fmt.Errorf("get active sheet index: %w", err)
		}
		workbook.SetActiveSheet(index)
	}

	output := &bytes.Buffer{}
	if err := workbook.Write(output); err != nil {
		return nil, fmt.Errorf("write workbook: %w", err)
	}
	return output, nil
}

func groupRecordsByMonth(records []outboundRecord) (map[string][]outboundRecord, []string) {
	if len(records) == 0 {
		return map[string][]outboundRecord{"出货列表": nil}, []string{"出货列表"}
	}

	recordsBySheet := make(map[string][]outboundRecord)
	firstSeen := make(map[string]int)
	for i, record := range records {
		sheetName := sheetNameFromDate(record.ApplyDate)
		if _, ok := recordsBySheet[sheetName]; !ok {
			firstSeen[sheetName] = i
		}
		recordsBySheet[sheetName] = append(recordsBySheet[sheetName], record)
	}

	sheetOrder := make([]string, 0, len(recordsBySheet))
	for sheetName := range recordsBySheet {
		sheetOrder = append(sheetOrder, sheetName)
	}
	sort.SliceStable(sheetOrder, func(i, j int) bool {
		return firstSeen[sheetOrder[i]] < firstSeen[sheetOrder[j]]
	})
	return recordsBySheet, sheetOrder
}

func writeOutboundSheet(workbook *excelize.File, sheetName string, records []outboundRecord) error {
	if err := writeOutboundHeader(workbook, sheetName); err != nil {
		return err
	}

	for rowIndex, record := range records {
		row := rowIndex + 6
		values := map[int]any{
			1: record.Customer,
			2: record.Applicant,
			3: record.DocumentNo,
			4: record.ContractNo,
			5: record.Contact,
			7: "是",
		}
		for col, value := range values {
			if err := setCell(workbook, sheetName, col, row, value); err != nil {
				return err
			}
		}
		for _, productName := range outputProducts {
			quantity := record.ProductCounts[productName]
			if quantity == 0 {
				continue
			}
			if err := setCell(workbook, sheetName, outputProductColumn[productName], row, numberValue(quantity)); err != nil {
				return err
			}
		}
	}

	return applyOutboundStyles(workbook, sheetName, len(records))
}

func buildOutboundLog(records []outboundRecord, mappingRows int) string {
	var builder strings.Builder
	now := time.Now().Format("2006-01-02 15:04:05")

	builder.WriteString("# 出库表转换日志\n\n")
	builder.WriteString(fmt.Sprintf("- 生成时间：%s\n", now))
	builder.WriteString("- 转换类型：出库表\n")
	builder.WriteString(fmt.Sprintf("- 硬件产品映射条数：%d\n", mappingRows))
	builder.WriteString(fmt.Sprintf("- 输出记录数：%d\n\n", len(records)))

	builder.WriteString("## 输出货品列\n\n")
	builder.WriteString("| 输出列 | 货品名称 |\n")
	builder.WriteString("| --- | --- |\n")
	for _, productName := range outputProducts {
		builder.WriteString(fmt.Sprintf("| %s | %s |\n", columnName(outputProductColumn[productName]), escapeMarkdown(productName)))
	}

	builder.WriteString("\n## 逐行转换明细\n\n")
	for _, record := range records {
		builder.WriteString(fmt.Sprintf("### 原文件第 %d 行 / 单据编号 %s\n\n", record.SourceRow, escapeMarkdown(record.DocumentNo)))
		builder.WriteString(fmt.Sprintf("- 输出 Sheet：%s\n", escapeMarkdown(record.SheetName)))
		builder.WriteString(fmt.Sprintf("- 输出行号：%d\n\n", record.OutputRow))
		builder.WriteString("| 输出字段 | 来源字段 | 写入值 |\n")
		builder.WriteString("| --- | --- | --- |\n")
		builder.WriteString(fmt.Sprintf("| 客户名称 | 客户名称 | %s |\n", escapeMarkdown(record.Customer)))
		builder.WriteString(fmt.Sprintf("| 提出人 | 申请人 | %s |\n", escapeMarkdown(record.Applicant)))
		builder.WriteString(fmt.Sprintf("| 单据编号 | 单据编号 | %s |\n", escapeMarkdown(record.DocumentNo)))
		builder.WriteString(fmt.Sprintf("| 合同编号 | 合同编号 | %s |\n", escapeMarkdown(record.ContractNo)))
		builder.WriteString(fmt.Sprintf("| 收货人 收货电话 | 收货人 + 收货电话 | %s |\n", escapeMarkdown(record.Contact)))
		builder.WriteString("| 发货时间 | 固定规则 |  |\n")
		builder.WriteString("| ERP流程是否流转到仓库方 | 固定规则 | 是 |\n\n")

		if len(record.Products) == 0 {
			builder.WriteString("未解析到任何产品明细。\n\n")
			continue
		}

		builder.WriteString("| 原产品名称 | 数量 | 映射货品名称 | 输出列 | 处理结果 | 说明 |\n")
		builder.WriteString("| --- | ---: | --- | --- | --- | --- |\n")
		for _, product := range record.Products {
			outputCol := ""
			if product.OutputCol > 0 {
				outputCol = product.OutputCell
			}
			builder.WriteString(fmt.Sprintf(
				"| %s | %s | %s | %s | %s | %s |\n",
				escapeMarkdown(firstNonEmpty(product.SourceName, product.Raw)),
				formatQuantity(product.Quantity),
				escapeMarkdown(product.GoodsName),
				outputCol,
				product.Status,
				escapeMarkdown(product.Description),
			))
		}

		if len(record.ProductCounts) > 0 {
			builder.WriteString("\n汇总写入：\n\n")
			builder.WriteString("| 输出单元格 | 货品名称 | 数量 |\n")
			builder.WriteString("| --- | --- | ---: |\n")
			for _, productName := range outputProducts {
				quantity := record.ProductCounts[productName]
				if quantity == 0 {
					continue
				}
				cell := fmt.Sprintf("%s%d", columnName(outputProductColumn[productName]), record.OutputRow)
				builder.WriteString(fmt.Sprintf("| %s | %s | %s |\n", cell, escapeMarkdown(productName), formatQuantity(quantity)))
			}
		}
		builder.WriteString("\n")
	}

	return builder.String()
}

func writeOutboundHeader(workbook *excelize.File, sheetName string) error {
	merges := [][2]string{
		{"A1", "AB1"},
		{"H3", "AA3"},
		{"H4", "I4"},
		{"J4", "P4"},
		{"Q4", "R4"},
		{"S4", "T4"},
		{"U4", "V4"},
	}
	for _, merge := range merges {
		if err := workbook.MergeCell(sheetName, merge[0], merge[1]); err != nil {
			return fmt.Errorf("merge %s:%s: %w", merge[0], merge[1], err)
		}
	}

	values := map[string]any{
		"A1":  " 章管家出库发货登记表",
		"A3":  "客户名称",
		"B3":  "提出人",
		"C3":  "单据编号",
		"D3":  "合同编号",
		"E3":  "收货人 收货电话",
		"F3":  "发货时间",
		"G3":  "ERP流程是否流转到仓库方",
		"H3":  "货品名称及数量",
		"AB3": "备注",
		"H4":  "用印工作台",
		"J4":  "智能印章",
		"Q4":  "印章柜",
		"S4":  "二代柜",
		"U4":  "文件柜",
		"W4":  "扫描仪",
		"X4":  "多功能一体机",
		"Y4":  "工作站mini（主机）",
		"Z4":  "工作站mini（格口）",
		"AA4": "工作站SC2C (小柜)",
	}
	for productName, col := range outputProductColumn {
		if col <= 22 {
			cell, err := excelize.CoordinatesToCellName(col, 5)
			if err != nil {
				return err
			}
			values[cell] = productName
		}
	}

	for cell, value := range values {
		if err := workbook.SetCellValue(sheetName, cell, value); err != nil {
			return fmt.Errorf("set %s: %w", cell, err)
		}
	}
	return nil
}

func applyOutboundStyles(workbook *excelize.File, sheetName string, recordCount int) error {
	titleStyle, err := workbook.NewStyle(&excelize.Style{
		Font:      &excelize.Font{Bold: true, Size: 16},
		Alignment: &excelize.Alignment{Horizontal: "center", Vertical: "center"},
	})
	if err != nil {
		return err
	}
	headerStyle, err := workbook.NewStyle(&excelize.Style{
		Font:      &excelize.Font{Bold: true},
		Alignment: &excelize.Alignment{Horizontal: "center", Vertical: "center", WrapText: true},
		Border:    tableBorders(),
	})
	if err != nil {
		return err
	}
	dataStyle, err := workbook.NewStyle(&excelize.Style{
		Alignment: &excelize.Alignment{Horizontal: "center", Vertical: "center", WrapText: true},
		Border:    tableBorders(),
	})
	if err != nil {
		return err
	}

	if err := workbook.SetCellStyle(sheetName, "A1", "AB1", titleStyle); err != nil {
		return err
	}
	if err := workbook.SetCellStyle(sheetName, "A3", "AB5", headerStyle); err != nil {
		return err
	}
	if recordCount > 0 {
		endCell, err := excelize.CoordinatesToCellName(28, recordCount+5)
		if err != nil {
			return err
		}
		if err := workbook.SetCellStyle(sheetName, "A6", endCell, dataStyle); err != nil {
			return err
		}
	}

	widths := map[string]float64{
		"A": 28, "B": 12, "C": 18, "D": 18, "E": 24, "F": 12, "G": 18,
		"H": 14, "I": 14, "J": 14, "K": 14, "L": 14, "M": 14, "N": 14, "O": 14, "P": 14,
		"Q": 12, "R": 12, "S": 12, "T": 12, "U": 12, "V": 12,
		"W": 12, "X": 16, "Y": 18, "Z": 18, "AA": 18, "AB": 18,
	}
	for col, width := range widths {
		if err := workbook.SetColWidth(sheetName, col, col, width); err != nil {
			return err
		}
	}
	for row := 1; row <= recordCount+5; row++ {
		height := 22.0
		if row == 1 {
			height = 30
		}
		if row >= 3 && row <= 5 {
			height = 28
		}
		if err := workbook.SetRowHeight(sheetName, row, height); err != nil {
			return err
		}
	}
	return nil
}

func tableBorders() []excelize.Border {
	return []excelize.Border{
		{Type: "left", Color: "000000", Style: 1},
		{Type: "top", Color: "000000", Style: 1},
		{Type: "right", Color: "000000", Style: 1},
		{Type: "bottom", Color: "000000", Style: 1},
	}
}

func sheetNameFromDate(value string) string {
	value = strings.TrimSpace(value)
	for _, layout := range []string{"2006-01-02", "2006/1/2", "2006/01/02", "2006-1-2"} {
		parsed, err := time.Parse(layout, value)
		if err == nil {
			return fmt.Sprintf("%d年%d月", parsed.Year(), int(parsed.Month()))
		}
	}
	return "出货列表"
}

func indexHeaders(row []string) headerIndex {
	headers := make(headerIndex, len(row))
	for i, value := range row {
		value = strings.TrimSpace(value)
		if value != "" {
			headers[value] = i
		}
	}
	return headers
}

func valueAt(row []string, index int) string {
	if index < 0 || index >= len(row) {
		return ""
	}
	return strings.TrimSpace(row[index])
}

func (r sourceRow) get(index int) string {
	return valueAt(r.values, index)
}

func cleanPhone(value string) string {
	value = strings.TrimSpace(value)
	return strings.TrimSuffix(value, ".0")
}

func joinNonEmpty(values ...string) string {
	parts := make([]string, 0, len(values))
	for _, value := range values {
		value = strings.TrimSpace(value)
		if value != "" {
			parts = append(parts, value)
		}
	}
	return strings.Join(parts, " ")
}

func numberValue(value float64) any {
	if math.Mod(value, 1) == 0 {
		return int(value)
	}
	return value
}

func setCell(workbook *excelize.File, sheetName string, col int, row int, value any) error {
	cell, err := excelize.CoordinatesToCellName(col, row)
	if err != nil {
		return err
	}
	return workbook.SetCellValue(sheetName, cell, value)
}

func columnName(col int) string {
	name, err := excelize.ColumnNumberToName(col)
	if err != nil {
		return ""
	}
	return name
}

func escapeMarkdown(value string) string {
	value = strings.ReplaceAll(value, "\n", " ")
	value = strings.ReplaceAll(value, "\r", " ")
	value = strings.ReplaceAll(value, "|", "\\|")
	return strings.TrimSpace(value)
}

func firstNonEmpty(values ...string) string {
	for _, value := range values {
		if strings.TrimSpace(value) != "" {
			return value
		}
	}
	return ""
}

func formatQuantity(value float64) string {
	if value == 0 {
		return ""
	}
	if math.Mod(value, 1) == 0 {
		return strconv.Itoa(int(value))
	}
	return strconv.FormatFloat(value, 'f', -1, 64)
}
