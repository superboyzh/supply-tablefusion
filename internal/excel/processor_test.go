package excel

import (
	"os"
	"path/filepath"
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
