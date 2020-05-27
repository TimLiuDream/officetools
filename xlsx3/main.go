package main

import (
	"fmt"
	"log"
	"math"

	"github.com/tealeg/xlsx"
)

func main() {
	func2()
}

func func1() {
	path := "/Users/tim/go/src/github.com/timliudream/GolangTraining/xlsx3/xlsx3test.xlsx"

	xlFile, err := xlsx.OpenFile(path)
	if err != nil {
		log.Fatal(err)
	}
	mergeCellValue := ""
	hMergeScope := make([]int, 2) // 列合并的范围，第一个值就是起始点，第二个值就是结束点
	hMergeScope[0] = math.MaxInt8
	vMergeScope := make([]int, 2) // 行合并的范围，第一个值就是起始点，第二个值就是结束点
	vMergeScope[0] = math.MaxInt8

	for _, sheet := range xlFile.Sheets {
		for rowIndex, row := range sheet.Rows {
			if rowIndex < vMergeScope[0] || rowIndex > vMergeScope[1] {
				mergeCellValue = ""
				vMergeScope[0] = math.MaxInt8
			}
			cells := row.Cells
			for cellIndex := 0; cellIndex < len(cells); cellIndex++ {
				if cellIndex >= vMergeScope[0] && cellIndex <= vMergeScope[1] {
					fmt.Printf("行：%d,列：%d，值:%s\n", rowIndex+1, cellIndex+1, mergeCellValue)
				} else {
					text := cells[cellIndex].String()

					h := cells[cellIndex].HMerge
					if h > 0 {
						mergeCellValue = text
						maxHMerge := cellIndex + h
						hMergeScope[0] = cellIndex
						hMergeScope[1] = maxHMerge
					}
					v := cells[cellIndex].VMerge
					if v > 0 {
						mergeCellValue = text
						maxVMerge := rowIndex + v
						vMergeScope[0] = rowIndex
						vMergeScope[1] = maxVMerge
						mergeCellValue = text
						maxHMerge := cellIndex + h
						hMergeScope[0] = cellIndex
						hMergeScope[1] = maxHMerge
					}

					fmt.Printf("行：%d,列：%d，值:%s\n", rowIndex+1, cellIndex+1, text)
				}
			}
		}
	}
}

func func2() {
	fmt.Println("sdf")
	path := "/Users/tim/Downloads/模板.xls"

	xlFile, err := xlsx.OpenFile(path)
	if err != nil {
		log.Fatal(err)
	}
	fmt.Println("sdf1")
	for _, sheet := range xlFile.Sheets {
		for rowIndex, row := range sheet.Rows {
			for cellIndex, cell := range row.Cells {
				text := cell.Value
				fmt.Printf("行：%d,列：%d，值:%s\n", rowIndex+1, cellIndex+1, text)
			}
		}
	}
}
