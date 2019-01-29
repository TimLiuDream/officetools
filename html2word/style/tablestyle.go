package style

import (
	"baliance.com/gooxml/color"
	"baliance.com/gooxml/document"
	"baliance.com/gooxml/measurement"
	"baliance.com/gooxml/schema/soo/wml"
	"github.com/timliudream/officetools/html2word/model"
)

// SetTable 往word写表格
func SetTable(rowCount, colCount int, tableCellMap map[string]*model.TableCell, cells []*model.TableCell) error {
	table := Doc.AddTable()
	table.Properties().SetWidthPercent(100)
	borders := table.Properties().Borders()
	borders.SetAll(wml.ST_BorderSingle, color.Auto, 2*measurement.Point)

	// 先创建rowCount、colCount的表格
	// 然后再根据tableCellMap来更改表格样式
	for rowIndex := 0; rowIndex < rowCount; rowIndex++ {
		row := table.AddRow()
		for {
			cellValue := cells[0]
			if cellValue.VMerge == 0 && cellValue.HMerge == 0 {
				c := row.AddCell()
				run := c.AddParagraph().AddRun()
				run.AddText(cellValue.Value)
			} else if cellValue.VMerge != 0 && cellValue.HMerge == 0 {
				c := row.AddCell()
				c.Properties().SetVerticalMerge(wml.ST_MergeRestart)
				c.AddParagraph().AddRun().AddText(cellValue.Value)
			} else if cellValue.HMerge != 0 && cellValue.VMerge == 0 {
				c := row.AddCell()
				c.Properties().SetColumnSpan(cellValue.HMerge)
				c.AddParagraph().AddRun().AddText(cellValue.Value)
			} else {
				c := row.AddCell()
				c.Properties().SetColumnSpan(2)
				c.Properties().SetVerticalMerge(wml.ST_MergeRestart)
				c.AddParagraph().AddRun().AddText(cellValue.Value)
			}
			cells = cells[1:]
			if len(cells) == 0 {
				break
			}
			if cells[0].RowIndex != rowIndex {
				break
			}
		}
	}

	//for _, cell := range cells {
	//	var row document.Row
	//	if cell.RowIndex+1 > len(table.Rows()) {
	//		row = table.AddRow()
	//	} else {
	//		row = table.Rows()[cell.RowIndex]
	//	}
	//	if cell.HMerge == 0 && cell.VMerge == 0 {
	//		c := row.AddCell()
	//		run := c.AddParagraph().AddRun()
	//		run.AddText(cell.Value)
	//	} else if cell.HMerge != 0 && cell.VMerge == 0 {
	//		c := row.AddCell()
	//		c.Properties().SetColumnSpan(cell.HMerge)
	//		run := c.AddParagraph().AddRun()
	//		run.AddText(cell.Value)
	//	} else if cell.HMerge == 0 && cell.VMerge != 0 {
	//		c := row.AddCell()
	//		c.Properties().SetVerticalMerge(wml.ST_MergeRestart)
	//		for ri := 1; ri < cell.VMerge; ri++ {
	//			r := table.AddRow()
	//			c := r.AddCell()
	//			c.Properties().SetVerticalMerge(wml.ST_MergeContinue)
	//		}
	//		run := c.AddParagraph().AddRun()
	//		run.AddText(cell.Value)
	//	} else {
	//		c := row.AddCell()
	//		c.Properties().SetColumnSpan(cell.HMerge)
	//		c.Properties().SetVerticalMerge(wml.ST_MergeRestart)
	//		for ri := 1; ri < cell.VMerge; ri++ {
	//			r := table.AddRow()
	//			c := r.AddCell()
	//			c.Properties().SetVerticalMerge(wml.ST_MergeContinue)
	//		}
	//		run := c.AddParagraph().AddRun()
	//		run.AddText(cell.Value)
	//	}
	//}

	//// 先创建rowCount、colCount的表格
	//// 然后再根据tableCellMap来更改表格样式
	//for rowIndex := 0; rowIndex < rowCount; rowIndex++ {
	//	row := table.AddRow()
	//	for colIndex := 0; colIndex < colCount; colIndex++ {
	//		cell := row.AddCell()
	//		run := cell.AddParagraph().AddRun()
	//		run.AddText("")
	//	}
	//}
	//
	//for _, cell := range cells {
	//	if cell.VMerge == 0 && cell.HMerge == 0 {
	//		rs := table.Rows()
	//		r := rs[cell.RowIndex]
	//		rcs := r.Cells()
	//		spanCount := 0
	//		ci := cell.ColIndex
	//		// 因前面格子有合并格子而导致列索引对不上的补救方法
	//		for i := 0; i < len(rcs); i++ {
	//			if rcs[i].X().TcPr != nil {
	//				if rcs[i].X().TcPr.GridSpan != nil {
	//					gridSpanInt64 := rcs[i].X().TcPr.GridSpan.ValAttr
	//					gridSpan := int(gridSpanInt64)
	//					if spanCount+gridSpan == cell.ColIndex {
	//						ci = i + 1
	//						break
	//					}
	//				}
	//			}
	//		}
	//		c := rcs[ci]
	//		run := c.AddParagraph().AddRun()
	//		run.AddText(cell.Value)
	//	} else if cell.VMerge != 0 && cell.HMerge == 0 {
	//		rs := table.Rows()
	//		r := rs[cell.RowIndex]
	//		rcs := r.Cells()
	//		c := rcs[cell.ColIndex]
	//		c.Properties().SetVerticalMerge(wml.ST_MergeRestart)
	//		for ri := 1; ri < cell.VMerge; ri++ {
	//			r := rs[cell.RowIndex+ri]
	//			c := r.Cells()[cell.ColIndex]
	//			c.Properties().SetVerticalMerge(wml.ST_MergeContinue)
	//		}
	//		run := c.AddParagraph().AddRun()
	//		run.AddText(cell.Value)
	//	} else if cell.VMerge == 0 && cell.HMerge != 0 {
	//		rs := table.Rows()
	//		r := rs[cell.RowIndex]
	//		rcs := r.Cells()
	//		c := rcs[cell.ColIndex]
	//		c.Properties().SetColumnSpan(cell.HMerge)
	//		run := c.AddParagraph().AddRun()
	//		run.AddText(cell.Value)
	//		rcs = append(rcs[:cell.ColIndex+1], rcs[cell.ColIndex+cell.HMerge:]...)
	//		fmt.Println()
	//	} else {
	//
	//	}
	//}
	return nil
}

// calCount 因前面格子有合并格子而导致列索引对不上的补救方法
func calCount(row document.Row) (count int) {
	rowCells := row.Cells()
	for i := 0; i < len(rowCells); i++ {
		if rowCells[i].X().TcPr != nil {
			if rowCells[i].X().TcPr.GridSpan != nil {
				gridSpanInt64 := rowCells[i].X().TcPr.GridSpan.ValAttr
				gridSpan := int(gridSpanInt64)
				count += gridSpan
			}
		}
	}
	return count
}
