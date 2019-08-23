package main

import (
	"fmt"
	"log"

	"github.com/tealeg/xlsx"
)

func main() {
	path := "/Users/tim/go/src/github.com/timliudream/officetools/xlsx6/xlsx6.xlsx"

	xlFile, err := xlsx.OpenFile(path)
	if err != nil {
		log.Fatal(err)
	}
	for _, sheet := range xlFile.Sheets {
		cell := sheet.Cell(0, 0)
		fmt.Println(cell)
	}
}
