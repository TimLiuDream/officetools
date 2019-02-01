package style

import (
	"baliance.com/gooxml/document"
)

// SetTable 往word写表格
//func SetTable(rowCount, colCount int, tableCellMap map[string]*model.TableCell, cells []*model.TableCell) error {
//	table := Doc.AddTable()
//	table.Properties().SetWidthPercent(100)
//	borders := table.Properties().Borders()
//	borders.SetAll(wml.ST_BorderSingle, color.Auto, 2*measurement.Point)
//
//	for _, cell := range cells {
//		var row document.Row
//		if cell.RowIndex+1 > len(table.Rows()) {
//			// 这里需要判断上一行是不是齐了
//			if cell.RowIndex != 0 {
//				preRow := table.Rows()[cell.RowIndex-1]
//				preRowCellCount := calCount(preRow)
//				if preRowCellCount != colCount {
//					c := preRow.AddCell()
//					c.Properties().SetColumnSpan(colCount - preRowCellCount)
//					c.Properties().SetVerticalMerge(wml.ST_MergeContinue)
//				}
//			}
//			row = table.AddRow()
//		} else {
//			row = table.Rows()[cell.RowIndex]
//		}
//		if cell.ColIndex > calCount(row) {
//			colSpanCount := cell.ColIndex - calCount(row)
//			c := row.AddCell()
//			c.Properties().SetColumnSpan(colSpanCount)
//			c.Properties().SetVerticalMerge(wml.ST_MergeContinue)
//		}
//		if cell.HMerge == 0 && cell.VMerge == 0 {
//			c := row.AddCell()
//			run := c.AddParagraph().AddRun()
//			run.AddText(cell.Value)
//		} else if cell.HMerge != 0 && cell.VMerge == 0 {
//			c := row.AddCell()
//			c.Properties().SetColumnSpan(cell.HMerge)
//			run := c.AddParagraph().AddRun()
//			run.AddText(cell.Value)
//		} else if cell.HMerge == 0 && cell.VMerge != 0 {
//			c := row.AddCell()
//			c.Properties().SetVerticalMerge(wml.ST_MergeRestart)
//			run := c.AddParagraph().AddRun()
//			run.AddText(cell.Value)
//		} else {
//			c := row.AddCell()
//			c.Properties().SetColumnSpan(cell.HMerge)
//			c.Properties().SetVerticalMerge(wml.ST_MergeRestart)
//			run := c.AddParagraph().AddRun()
//			run.AddText(cell.Value)
//		}
//	}
//	// 再检查一下最后一行的格子是不是齐了
//	lastRow := table.Rows()[len(table.Rows())-1]
//	lastRowCellCount := calCount(lastRow)
//	if lastRowCellCount != colCount {
//		c := lastRow.AddCell()
//		c.Properties().SetColumnSpan(colCount - lastRowCellCount)
//		c.Properties().SetVerticalMerge(wml.ST_MergeContinue)
//	}
//
//	return nil
//}

// calCount 因前面格子有合并格子而导致列索引对不上的补救方法
func calCount(row document.Row) (count int) {
	rowCells := row.Cells()
	for i := 0; i < len(rowCells); i++ {
		if rowCells[i].X().TcPr != nil {
			if rowCells[i].X().TcPr.GridSpan != nil {
				gridSpanInt64 := rowCells[i].X().TcPr.GridSpan.ValAttr
				gridSpan := int(gridSpanInt64)
				count += gridSpan
				continue
			}
		}
		count++
	}
	return count
}
