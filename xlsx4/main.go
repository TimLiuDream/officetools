package main

import (
	"fmt"
	"log"
	"strconv"

	"github.com/tealeg/xlsx"
)

type MergeCellValue struct {
	Row   int
	Col   int
	Value string
}

func main() {
	// 此demo用于拆分excel的单元格
	// 1. 遍历所有格子
	// 2. 如果格子不是合并的，那就普通操作
	// 3. 如果格子是合并的，那就拆分，并算出各个格子的坐标，并把值写进对象MergeCellValue中
	mcvMap := make(map[string]MergeCellValue)
	path := "/Users/tim/go/src/github.com/timliudream/GolangTraining/xlsx4/xlsx4test.xlsx"

	xlFile, err := xlsx.OpenFile(path)
	if err != nil {
		log.Fatal(err)
	}

	for _, sheet := range xlFile.Sheets {
		for rowIndex, row := range sheet.Rows {
			cells := row.Cells
			for cellIndex, cell := range cells {
				if cell.HMerge == 0 && cell.VMerge == 0 {
					key := getCellKey(rowIndex, cellIndex)
					value, ok := mcvMap[key]
					if !ok {
						fmt.Printf("行：%d,列：%d，值:%s\n", rowIndex+1, cellIndex+1, cell.String())
					} else {
						fmt.Printf("行：%d,列：%d，值:%s\n", value.Row+1, value.Col+1, value.Value)
					}
				} else {
					fmt.Printf("行：%d,列：%d，值:%s\n", rowIndex+1, cellIndex+1, cell.String())

					hMerge := cell.HMerge
					vMerge := cell.VMerge

					if hMerge != 0 && vMerge == 0 {
						for index := 1; index <= hMerge; index++ {
							mcv := MergeCellValue{
								Row:   rowIndex,
								Col:   cellIndex + index,
								Value: cell.String(),
							}
							key := getCellKey(rowIndex, cellIndex+index)
							mcvMap[key] = mcv
						}
					} else if vMerge != 0 && hMerge == 0 {
						for index := 1; index <= vMerge; index++ {
							mcv := MergeCellValue{
								Row:   rowIndex + index,
								Col:   cellIndex,
								Value: cell.String(),
							}
							key := getCellKey(rowIndex+index, cellIndex)
							mcvMap[key] = mcv
						}
					} else {
						for vIndex := 0; vIndex <= vMerge; vIndex++ {
							for hIndex := 0; hIndex <= hMerge; hIndex++ {
								mcv := MergeCellValue{
									Row:   rowIndex + vIndex,
									Col:   cellIndex + hIndex,
									Value: cell.String(),
								}
								key := getCellKey(rowIndex+vIndex, cellIndex+hIndex)
								mcvMap[key] = mcv
							}
						}
					}
				}
			}
		}
	}
}

func getCellKey(rowIndex, colIndex int) string {
	return strconv.Itoa(rowIndex) + "," + strconv.Itoa(colIndex)
}
