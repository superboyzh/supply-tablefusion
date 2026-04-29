package excel

import (
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/xuri/excelize/v2"
)

func TestTransformOutboundXLS(t *testing.T) {
	input, err := os.Open(filepath.Join("..", "..", "示例文件", "原文件", "原文件_出货列表.xls"))
	if err != nil {
		t.Fatal(err)
	}
	defer input.Close()

	mapping, err := os.Open(filepath.Join("..", "..", "示例文件", "硬件产品信息.xlsx"))
	if err != nil {
		t.Fatal(err)
	}
	defer mapping.Close()

	output, err := Transform(input, SourceTypeOutbound, mapping)
	if err != nil {
		t.Fatal(err)
	}
	if output.LogMarkdown == "" {
		t.Fatal("expected markdown log")
	}

	workbook, err := excelize.OpenReader(output.Workbook)
	if err != nil {
		t.Fatal(err)
	}
	defer workbook.Close()

	const sheet = "2026年4月"
	if index, err := workbook.GetSheetIndex(sheet); err != nil || index == -1 {
		t.Fatalf("expected sheet %q, index=%d, err=%v", sheet, index, err)
	}

	assertCell(t, workbook, sheet, "A3", "客户名称")
	assertCell(t, workbook, sheet, "H3", "货品名称及数量")
	assertCell(t, workbook, sheet, "M5", "智能终端S2P(光敏)")
	assertCell(t, workbook, sheet, "O5", "章管家 MS1(光敏)")

	row := findRowByDocumentNo(t, workbook, sheet, "CH202604021345")
	assertCell(t, workbook, sheet, cellName(t, 1, row), "友邦人寿保险有限公司")
	assertCell(t, workbook, sheet, cellName(t, 2, row), "滕齐华")
	assertCell(t, workbook, sheet, cellName(t, 4, row), "2025-XWJ-12082")
	assertCell(t, workbook, sheet, cellName(t, 5, row), "陈泽汉 13976030955")
	assertCell(t, workbook, sheet, cellName(t, 6, row), "")
	assertCell(t, workbook, sheet, cellName(t, 7, row), "是")
	assertCell(t, workbook, sheet, cellName(t, 13, row), "3")
	assertCell(t, workbook, sheet, cellName(t, 15, row), "10")

	row = findRowByDocumentNo(t, workbook, sheet, "CH202604276874")
	assertCell(t, workbook, sheet, cellName(t, 1, row), "河南永坤水利建筑工程有限公司")
	assertCell(t, workbook, sheet, cellName(t, 4, row), "2026-XWJ-04140")
	assertCell(t, workbook, sheet, cellName(t, 5, row), "宋艳伟 15539763983")
	assertCell(t, workbook, sheet, cellName(t, 11, row), "2")

	row = findRowByDocumentNo(t, workbook, sheet, "CH202604275723")
	assertCell(t, workbook, sheet, cellName(t, 1, row), "山西知闻品牌管理有限公司")
	assertCell(t, workbook, sheet, cellName(t, 4, row), "2026-XWJ-04075")
	assertCell(t, workbook, sheet, cellName(t, 5, row), "尹雪蓉 13007085719")
	assertCell(t, workbook, sheet, cellName(t, 11, row), "1")

	row = findRowByDocumentNo(t, workbook, sheet, "CH202604273891")
	assertCell(t, workbook, sheet, cellName(t, 1, row), "上海建业信息科技股份有限公司")
	assertCell(t, workbook, sheet, cellName(t, 4, row), "2026-XWJ-04133")
	assertCell(t, workbook, sheet, cellName(t, 8, row), "1")
	assertCell(t, workbook, sheet, cellName(t, 13, row), "3")
}

func TestTransformWeidianXLSX(t *testing.T) {
	input, err := os.Open(filepath.Join("..", "..", "示例文件", "原文件", "原文件_微店.xlsx"))
	if err != nil {
		t.Fatal(err)
	}
	defer input.Close()

	output, err := Transform(input, SourceTypeWeidian, nil)
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(output.LogMarkdown, "订单状态为“已关闭”的行不输出") {
		t.Fatal("expected weidian markdown log")
	}
	if strings.Contains(output.LogMarkdown, "844217256880654") {
		t.Fatal("closed order should not appear in log")
	}

	workbook, err := excelize.OpenReader(output.Workbook)
	if err != nil {
		t.Fatal(err)
	}
	defer workbook.Close()

	const sheet = "2026年4月"
	if index, err := workbook.GetSheetIndex(sheet); err != nil || index == -1 {
		t.Fatalf("expected sheet %q, index=%d, err=%v", sheet, index, err)
	}

	assertCell(t, workbook, sheet, "A2", "客户名称")
	assertCell(t, workbook, sheet, "E2", "配件名称及数量")
	assertCell(t, workbook, sheet, "H3", "定位卡片")
	assertCell(t, workbook, sheet, "O3", "3M胶")
	assertCell(t, workbook, sheet, "P3", "环形胶")
	assertCell(t, workbook, sheet, "Q3", "定制章环")
	assertCell(t, workbook, sheet, "R3", "定制光敏底座")
	assertCell(t, workbook, sheet, "X3", "印章电源")
	assertCell(t, workbook, sheet, "Y3", "光敏配件一套")
	assertCell(t, workbook, sheet, "Z3", "铜章配件一套")
	assertCell(t, workbook, sheet, "AB3", "铜章快拆配件")

	row := findRowByApplicant(t, workbook, sheet, "钟小姐")
	assertCell(t, workbook, sheet, cellName(t, 1, row), "微店")
	assertCell(t, workbook, sheet, cellName(t, 4, row), "13620564479广东省 深圳市 南山区 沙河街道 侨香路智慧广场A1座1201")
	assertCell(t, workbook, sheet, cellName(t, 15, row), "10")
	assertCell(t, workbook, sheet, cellName(t, 16, row), "20")
	assertCell(t, workbook, sheet, cellName(t, 30, row), "")

	row = findRowByApplicant(t, workbook, sheet, "谢女士")
	assertCell(t, workbook, sheet, cellName(t, 30, row), "是")

	row = findRowByApplicant(t, workbook, sheet, "王雅茜")
	assertCell(t, workbook, sheet, cellName(t, 7, row), "2")

	row = findRowByApplicant(t, workbook, sheet, "杜晓永")
	assertCell(t, workbook, sheet, cellName(t, 5, row), "2")

	row = findRowByApplicant(t, workbook, sheet, "木木")
	assertCell(t, workbook, sheet, cellName(t, 5, row), "10")
	assertCell(t, workbook, sheet, cellName(t, 6, row), "10")

	row = findRowByApplicant(t, workbook, sheet, "朱洪")
	assertCell(t, workbook, sheet, cellName(t, 6, row), "1")
	assertCell(t, workbook, sheet, cellName(t, 15, row), "1")

	row = findRowByApplicant(t, workbook, sheet, "薛坤")
	assertCell(t, workbook, sheet, cellName(t, 17, row), "1")
	assertCell(t, workbook, sheet, cellName(t, 25, row), "1")
	assertCell(t, workbook, sheet, cellName(t, 31, row), "")

	row = findRowByApplicant(t, workbook, sheet, "郭昊")
	assertCell(t, workbook, sheet, cellName(t, 24, row), "2")

	row = findRowByApplicant(t, workbook, sheet, "邹君毫")
	assertCell(t, workbook, sheet, cellName(t, 23, row), "1")

	row = findRowByApplicant(t, workbook, sheet, "赵一恒")
	assertCell(t, workbook, sheet, cellName(t, 8, row), "3")
	assertCell(t, workbook, sheet, cellName(t, 24, row), "1")

	row = findRowByApplicant(t, workbook, sheet, "温雅")
	assertCell(t, workbook, sheet, cellName(t, 26, row), "1")

	row = findRowByApplicant(t, workbook, sheet, "严丽彬")
	assertCell(t, workbook, sheet, cellName(t, 25, row), "1")
	assertCell(t, workbook, sheet, cellName(t, 28, row), "1")
}

func TestParseWeidianPartsCustomPhotosensitiveBase(t *testing.T) {
	counts, entries, notes := parseWeidianParts("6239519854", "2", "定制光敏底座")
	if len(notes) != 0 {
		t.Fatalf("expected no notes, got %v", notes)
	}
	if len(entries) != 1 || entries[0].PartName != "定制光敏底座" || entries[0].OutputCol != 18 {
		t.Fatalf("unexpected entries: %+v", entries)
	}
	if counts["定制光敏底座"] != 2 {
		t.Fatalf("expected 定制光敏底座 count 2, got %v", counts["定制光敏底座"])
	}
}

func TestParseWeidianPartsPhotosensitiveBase(t *testing.T) {
	counts, entries, notes := parseWeidianParts("4294778913", "3", "光敏底座")
	if len(notes) != 0 {
		t.Fatalf("expected no notes, got %v", notes)
	}
	if len(entries) != 1 || entries[0].PartName != "光敏底座" || entries[0].OutputCol != 12 {
		t.Fatalf("unexpected entries: %+v", entries)
	}
	if counts["光敏底座"] != 3 {
		t.Fatalf("expected 光敏底座 count 3, got %v", counts["光敏底座"])
	}
}

func TestParseWeidianPartsUnmappedProductRemark(t *testing.T) {
	counts, entries, notes := parseWeidianParts("9999999999", "1", "旋转柜内置充电线*1m")
	if len(counts) != 0 {
		t.Fatalf("expected no mapped counts, got %v", counts)
	}
	if len(entries) != 1 || entries[0].Status != "未映射" {
		t.Fatalf("unexpected entries: %+v", entries)
	}
	if len(notes) != 1 || notes[0] != "旋转柜内置充电线*1m *1" {
		t.Fatalf("unexpected notes: %v", notes)
	}
}

func findRowByDocumentNo(t *testing.T, workbook *excelize.File, sheet string, documentNo string) int {
	t.Helper()

	rows, err := workbook.GetRows(sheet)
	if err != nil {
		t.Fatal(err)
	}
	for rowIndex, row := range rows {
		if len(row) >= 3 && row[2] == documentNo {
			return rowIndex + 1
		}
	}
	t.Fatalf("document %s not found", documentNo)
	return 0
}

func findRowByApplicant(t *testing.T, workbook *excelize.File, sheet string, applicant string) int {
	t.Helper()

	rows, err := workbook.GetRows(sheet)
	if err != nil {
		t.Fatal(err)
	}
	for rowIndex, row := range rows {
		if len(row) >= 2 && row[1] == applicant {
			return rowIndex + 1
		}
	}
	t.Fatalf("applicant %s not found", applicant)
	return 0
}

func assertCell(t *testing.T, workbook *excelize.File, sheet string, cell string, expected string) {
	t.Helper()

	actual, err := workbook.GetCellValue(sheet, cell)
	if err != nil {
		t.Fatal(err)
	}
	if actual != expected {
		t.Fatalf("%s!%s = %q, want %q", sheet, cell, actual, expected)
	}
}

func cellName(t *testing.T, col int, row int) string {
	t.Helper()

	cell, err := excelize.CoordinatesToCellName(col, row)
	if err != nil {
		t.Fatal(err)
	}
	return cell
}
