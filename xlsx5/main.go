package main

import (
	"log"

	"github.com/tealeg/xlsx"
)

// 让单元格的字体有黑色和红色
func main() {
	col := "用例名称"

	style := xlsx.NewStyle()
	font := *xlsx.NewFont(12, "DengXian")
	font.Family = 2
	font.Charset = 134
	font.Color = "#ff0000"
	font.Bold = false
	font.Italic = false
	font.Underline = false
	style.Font = font
	style.ApplyFont = true

	file := xlsx.NewFile()
	sheet, err := file.AddSheet("sheet1")
	if err != nil {
		log.Fatalln(err)
	}

	headRow := sheet.AddRow()
	cell := headRow.AddCell()
	cell.Value = col
	cell.SetStyle(style)
	err = file.Save("/Users/tim/go/src/github.com/timliudream/officetools/xlsx5/test.xlsx")
	if err != nil {
		log.Fatalln(err)
	}
}
