package main

import (
	"fmt"

	"github.com/tealeg/xlsx"
)

func main() {
	f, err := xlsx.OpenFile("/Users/tim/go/src/github.com/timliudream/officetools/xlsx7/1.xlsx")
	if err != nil {
		panic(err)
	}
	for _, sheet := range f.Sheets {
		for _, row := range sheet.Rows {
			for _, cell := range row.Cells {
				fmt.Println(cell)
			}
		}
	}
}
