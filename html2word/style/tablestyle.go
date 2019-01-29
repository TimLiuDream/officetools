package style

import (
	"baliance.com/gooxml/color"
	"baliance.com/gooxml/measurement"
	"baliance.com/gooxml/schema/soo/wml"
	"github.com/timliudream/officetools/html2word/model"
	"github.com/timliudream/officetools/html2word/utils"
)

// SetTable 往word写表格
func SetTable(rowCount, colCount int, tableCellMap map[string]*model.TableCell) error {
	table := Doc.AddTable()
	table.Properties().SetWidthPercent(100)
	borders := table.Properties().Borders()
	borders.SetAll(wml.ST_BorderSingle, color.Auto, 2*measurement.Point)

	for rowIndex := 0; rowIndex < rowCount; rowIndex++ {
		//row := table.AddRow()
		for colIndex := 0; colIndex < colCount; colIndex++ {
			cellKey := utils.GetCellKey(rowIndex, colIndex)
			_, ok := tableCellMap[cellKey]
			if !ok {
				//cellRun := row.AddCell().AddParagraph().AddRun()
				//cellRun.AddText(cellMap[cellKey])
			} else {
				//rowStart := mergeCellScope.RowScope.Start
				//rowEnd := mergeCellScope.RowScope.End
				//colStart := mergeCellScope.ColScope.Start
				//colEnd := mergeCellScope.ColScope.End
				//cell := row.AddCell()
				//cell.Properties().SetColumnSpan(colEnd - colStart + 1)
				//run := cell.AddParagraph().AddRun()
				//run.AddText(mergeCellScope.Value)
			}
		}
	}
	return nil
}
