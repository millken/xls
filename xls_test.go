package xls

import (
	"fmt"
	"strings"
	"testing"
)

func TestOpen(t *testing.T) {
	if xlFile, err := Open("testdata/float.xls", "utf-8"); err == nil {
		if sheet1 := xlFile.GetSheet(0); sheet1 != nil {
			fmt.Println("Total Lines ", sheet1.MaxRow, sheet1.Name)
			for i := 0; i <= 5; i++ {
				fmt.Printf("row %v point %v \n", i, sheet1.Row(i))
				if sheet1.Row(i) == nil {
					continue
				}
				row := sheet1.Row(i)
				for index := row.FirstCol(); index < row.LastCol(); index++ {
					fmt.Println(index, "==>", row.Col(index), " ")
				}
			}
		}
	}
}

func TestMargin(t *testing.T) {
	xlFile, err := Open("testdata/sh_margin.xls", "UTF-8")
	if err != nil {
		t.Fatal(err)
	}

	fmt.Printf("Number of sheets: %d\n\n", xlFile.NumSheets())

	// 获取 Sheet 1 明细信息
	sheet := xlFile.GetSheet(1)
	if sheet == nil {
		t.Error("Sheet 1 is nil")
		return
	}

	fmt.Printf("Sheet: %s, MaxRow: %d\n\n", sheet.Name, sheet.MaxRow)

	// 统计有效行数
	validCount := 0
	aStockCount := 0

	for j := 0; j <= int(sheet.MaxRow); j++ {
		row := sheet.Row(j)
		if row == nil {
			continue
		}

		code := strings.TrimSpace(row.Col(0))
		name := strings.TrimSpace(row.Col(1))

		if code == "" || strings.Contains(code, "证券") || strings.Contains(code, "标的") {
			continue
		}

		validCount++

		// 检查是否为A股代码 (6开头)
		if len(code) == 6 && (strings.HasPrefix(code, "6") || strings.HasPrefix(code, "0") || strings.HasPrefix(code, "3")) {
			aStockCount++
		}

		// 打印前10行和每200行
		if j < 10 || j%200 == 0 {
			fmt.Printf("Row %d: code=%s, name=%s\n", j, code, name)
		}
	}

	fmt.Printf("\nTotal valid rows: %d\n", validCount)
	fmt.Printf("A-Stock codes: %d\n", aStockCount)

	// 查找600179
	fmt.Println("\n--- Searching for 600179 ---")
	for j := 0; j <= int(sheet.MaxRow); j++ {
		row := sheet.Row(j)
		if row == nil {
			continue
		}
		code := strings.TrimSpace(row.Col(0))
		if code == "600179" {
			fmt.Printf("Found 600179 at row %d: %+v\n", j, row.Col(1))
			return
		}
	}
	fmt.Println("600179 NOT FOUND!")
}
