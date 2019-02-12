package main

import (
	"baliance.com/gooxml/document"
	"baliance.com/gooxml/measurement"
	"baliance.com/gooxml/schema/soo/wml"
	"fmt"
	"log"
)

func main() {
	doc := document.New()
	// 无序列表
	notSortList(doc, "./toc/notsortlist.docx")

	doc1 := document.New()
	// 有序列表
	sortList(doc1, "./toc/sortlist.docx")
}

// notSortList 创建无序列表
func notSortList(doc *document.Document, path string) {
	nd := doc.Numbering.Definitions()[0]
	for i := 1; i < 5; i++ {
		p := doc.AddParagraph()
		p.SetNumberingLevel(i - 1)
		p.SetNumberingDefinition(nd)
		run := p.AddRun()
		run.AddText(fmt.Sprintf("Level %d", i))
	}
	err := doc.SaveToFile(path)
	if err != nil {
		log.Fatalln(err)
	}
}

// sortList 有序列表
func sortList(doc *document.Document, path string) {
	nd := doc.Numbering.AddDefinition()
	nd.SetMultiLevelType(wml.ST_MultiLevelTypeHybridMultilevel)
	// 新建列表系列
	for i := 1; i < 10; i++ {
		lvl := nd.AddLevel()
		lvl.SetFormat(wml.ST_NumberFormatCustom)
		lvl.SetAlignment(wml.ST_JcLeft)
		txt := fmt.Sprintf("%%%d.", i)
		lvl.SetText(txt)
		if i%2 == 0 {
			lvl.SetFormat(wml.ST_NumberFormatCustom)
			txt = fmt.Sprintf("%%%d.", i)
			lvl.SetText(txt)
		} else if i%3 == 0 {
			lvl.SetFormat(wml.ST_NumberFormatCustom)
			txt = fmt.Sprintf("%%%d.", i)
			lvl.SetText(txt)
		}
		lvl.Properties().SetLeftIndent(0.5 * measurement.Distance(i) * measurement.Inch)
	}

	// and finally paragraphs at different heading levels
	for i := 0; i < 4; i++ {
		para := doc.AddParagraph()
		para.SetNumberingDefinition(nd)
		para.SetNumberingLevel(0)
		para.AddRun().AddText("First Level")

		for i := 0; i < 3; i++ {
			para := doc.AddParagraph()
			para.SetNumberingDefinition(nd)
			para.SetNumberingLevel(1)
			para.AddRun().AddText("Second Level")

			para = doc.AddParagraph()
			para.SetNumberingDefinition(nd)
			para.SetNumberingLevel(2)
			para.AddRun().AddText("Third Level")
		}
	}

	err := doc.SaveToFile(path)
	if err != nil {
		log.Fatalln(err)
	}
}
