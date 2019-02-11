package wordstyle

import (
	"baliance.com/gooxml/color"
	"baliance.com/gooxml/measurement"
	"baliance.com/gooxml/schema/soo/wml"
	"github.com/timliudream/officetools/html2word/model"
)

func SetTable(vTable [][]*model.TableCell) {
	table := Doc.AddTable()
	table.Properties().SetWidthPercent(100)
	borders := table.Properties().Borders()
	borders.SetAll(wml.ST_BorderSingle, color.Auto, 2*measurement.Point)

	for _, rowCell := range vTable {
		// 新建一行
		row := table.AddRow()
		for i := 0; i < len(rowCell); {
			if rowCell[i].IsVMergeStart == false && rowCell[i].IsVMerge == false && rowCell[i].HMerge == 0 {
				c := row.AddCell()
				run := c.AddParagraph().AddRun()
				run.AddText(rowCell[i].Value)
				i++
			} else if rowCell[i].IsVMergeStart == true && rowCell[i].IsVMerge == true { // 有竖向合并,并且是开头
				if rowCell[i].HMerge > 1 { //有横向合并
					c := row.AddCell()
					c.Properties().SetColumnSpan(rowCell[i].HMerge)
					c.Properties().SetVerticalMerge(wml.ST_MergeRestart)
					run := c.AddParagraph().AddRun()
					run.AddText(rowCell[i].Value)
					i += rowCell[i].HMerge
				} else { // 没有横向合并
					c := row.AddCell()
					c.Properties().SetVerticalMerge(wml.ST_MergeRestart)
					run := c.AddParagraph().AddRun()
					run.AddText(rowCell[i].Value)
					i++
				}
			} else if rowCell[i].IsVMergeStart == false && rowCell[i].IsVMerge == true { // 有竖向合并，不是开头
				if rowCell[i].HMerge > 1 { //有横向合并
					c := row.AddCell()
					c.Properties().SetColumnSpan(rowCell[i].HMerge)
					c.Properties().SetVerticalMerge(wml.ST_MergeContinue)
					run := c.AddParagraph().AddRun()
					run.AddText(rowCell[i].Value)
					i += rowCell[i].HMerge
				} else { // 没有横向合并
					c := row.AddCell()
					c.Properties().SetVerticalMerge(wml.ST_MergeContinue)
					run := c.AddParagraph().AddRun()
					run.AddText(rowCell[i].Value)
					i++
				}
			} else {
				c := row.AddCell()
				c.Properties().SetColumnSpan(rowCell[i].HMerge)
				run := c.AddParagraph().AddRun()
				run.AddText(rowCell[i].Value)
				i += rowCell[i].HMerge
			}
		}
	}
}
