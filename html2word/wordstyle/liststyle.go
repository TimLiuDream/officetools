package wordstyle

import (
	"baliance.com/gooxml/document"
	"baliance.com/gooxml/measurement"
	"baliance.com/gooxml/schema/soo/wml"
	"fmt"
	"github.com/timliudream/officetools/html2word/model"
)

// SetNotSortList 往word写无序列表
func SetNotSortList(list []*model.NotSortItem, level int) {
	nd := Doc.Numbering.Definitions()[0]
	for _, item := range list {
		p := Doc.AddParagraph()
		p.SetNumberingLevel(level)
		p.SetNumberingDefinition(nd)
		run := p.AddRun()
		run.AddText(item.Value)
		l := level
		if len(item.NotSortItemList) > 0 {
			l++
			SetNotSortList(item.NotSortItemList, l)
		}
	}
}

// SetSortList 往word写有序列表
func SetSortList(list []*model.SortItem, level int) {
	nd := setSortListCode()
	for _, item := range list {
		p := Doc.AddParagraph()
		p.SetNumberingLevel(level)
		p.SetNumberingDefinition(nd)
		run := p.AddRun()
		run.AddText(item.Value)
		l := level
		if len(item.SortItemList) > 0 {
			l++
			SetSortList(item.SortItemList, l)
		}
	}
}

// setSortListCode 创建有序列表的编码
func setSortListCode() (nd document.NumberingDefinition) {
	nd = Doc.Numbering.AddDefinition()
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
	return nd
}
