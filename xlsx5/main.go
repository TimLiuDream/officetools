package main

import (
	"log"

	"github.com/tealeg/xlsx"
)

// 让单元格的字体有黑色和红色
func main() {
	col := "用例名称"
	required := "*"

	style := xlsx.NewStyle()
	font := xlsx.NewFont(12, "DengXian")
	font.Family = 2
	font.Charset = 134
	font.Color = "FF000000"
	font.Bold = false
	font.Italic = false
	font.Underline = false
	file := xlsx.NewFile()
	sheet, err := file.AddSheet("sheet1")
	if err != nil {
		log.Fatalln(err)
	}

	headRow := sheet.AddRow()
	cell := headRow.AddCell()
	cell.SetStyle(style)
	cell.Value = col + required
	err = file.Save("/Users/tim/go/src/github.com/timliudream/officetools/xlsx5/test.xlsx")
	if err != nil {
		log.Fatalln(err)
	}
}
