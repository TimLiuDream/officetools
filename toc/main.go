package main

import (
	"baliance.com/gooxml/document"
	"baliance.com/gooxml/measurement"
	"baliance.com/gooxml/schema/soo/wml"
)

func main() {
	doc := document.New()

	// Force the TOC to update upon opening the document
	//doc.Settings.SetUpdateFieldsOnOpen(true)

	nd := doc.Numbering.AddDefinition()
	for i := 0; i < 9; i++ {
		lvl := nd.AddLevel()
		//lvl.SetFormat(wml.ST_NumberFormatBullet)
		lvl.SetFormat(wml.ST_NumberFormatCustom)
		lvl.SetAlignment(wml.ST_JcEnd)
		lvl.Properties().SetLeftIndent(0.5 * measurement.Distance(i) * measurement.Inch)
	}

	// and finally paragraphs at different heading levels
	for i := 0; i < 4; i++ {
		para := doc.AddParagraph()
		para.SetNumberingDefinition(nd)
		para.SetNumberingLevel(1)
		para.AddRun().AddText("First Level")

		for i := 0; i < 3; i++ {
			para := doc.AddParagraph()
			para.SetNumberingDefinition(nd)
			para.SetNumberingLevel(2)
			para.AddRun().AddText("Second Level")

			para = doc.AddParagraph()
			para.SetNumberingDefinition(nd)
			para.SetNumberingLevel(3)
			para.AddRun().AddText("Third Level")
		}
	}
	doc.SaveToFile("toc.docx")
}