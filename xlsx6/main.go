package main

import (
	"fmt"
	"log"

	"github.com/tealeg/xlsx"
)

func main() {
	fill := xlsx.NewFill("", "FF000000", "#FFFFFF")
	style := xlsx.NewStyle()
	style.Fill = *fill
	font := *xlsx.NewFont(20, "Verdana")
	border := *xlsx.NewBorder("thin", "thin", "thin", "thin")
	alignment := xlsx.Alignment{Vertical: "center", WrapText: true, ShrinkToFit: true}
	style.Font = font
	// font.Color = "FF000000"
	style.Border = border
	style.Alignment = alignment
	style.ApplyFont = true
	style.ApplyBorder = true
	style.ApplyAlignment = true

	path := "/Users/tim/go/src/github.com/timliudream/officetools/xlsx6/testcase_import_template.xlsx"

	xlFile, err := xlsx.OpenFile(path)
	if err != nil {
		log.Fatal(err)
	}
	for _, sheet := range xlFile.Sheets {
		cell := sheet.Cell(0, 0)
		fmt.Println(cell)
	}
	sheet := xlFile.Sheets[0]
	for rIndex, _ := range sheet.Rows {
		for cIndex, _ := range sheet.Cols {
			cell := sheet.Cell(rIndex, cIndex)
			cell.SetStyle(style)
		}
	}

	cell := sheet.Cell(5, 5)
	cell.SetStyle(style)
	cell.SetString("123")
	xlFile.Save(path)
}
