package main

import (
	"baliance.com/gooxml/color"
	"baliance.com/gooxml/document"
	"baliance.com/gooxml/measurement"
	"baliance.com/gooxml/schema/soo/wml"
	"log"
)

func main() {
	doc := document.New()
	{
		table := doc.AddTable()
		// 4 inches wide
		table.Properties().SetWidthPercent(100)
		borders := table.Properties().Borders()
		// thin borders
		borders.SetAll(wml.ST_BorderSingle, color.Auto, measurement.Zero)


		row:=table.AddRow()
		cell:=row.AddCell()
		cell.Properties().SetColumnSpan(2)
		cell.Properties().SetVerticalMerge(wml.ST_MergeRestart)
		cell.AddParagraph().AddRun().AddText("行列合并")
		row.AddCell().AddParagraph().AddRun().AddText("1行3列")

		cell = row.AddCell()
		cell.Properties().SetVerticalMerge(wml.ST_MergeRestart)
		cell.AddParagraph().AddRun().AddText("行合并")

		row = table.AddRow()
		cell=row.AddCell()
		cell.Properties().SetColumnSpan(2)
		cell.Properties().SetVerticalMerge(wml.ST_MergeContinue)
		cell.AddParagraph().AddRun().AddText("行列合并1")
		row.AddCell().AddParagraph().AddRun().AddText("2行3列")

		cell = row.AddCell()
		cell.Properties().SetVerticalMerge(wml.ST_MergeContinue)
		cell.AddParagraph().AddRun().AddText("行合并")
	}

	if err := doc.Validate(); err != nil {
		log.Fatalf("error during validation: %s", err)
	}
	doc.SaveToFile("tables.docx")
}