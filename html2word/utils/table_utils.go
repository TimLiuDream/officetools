package utils

import (
	"github.com/PuerkitoBio/goquery"
	"github.com/timliudream/officetools/html2word/model"
	"golang.org/x/net/html"
	"log"
	"strconv"
)

// 计算表格的行列数
func CalTableRowColCount(s *goquery.Selection) (rowCount, colCount int) {
	rowSelection := s.Find("tbody tr")
	rowCount = len(rowSelection.Nodes)
	rowSelection.Each(func(i int, selection *goquery.Selection) {
		if i == 0 {
			cellSelection := selection.Find("td")
			for _, node := range cellSelection.Nodes {
				if node.Attr == nil {
					colCount += 1
				} else {
					has, colSpanCount := IsCellHasColSpanAttr(node)
					if has {
						colCount += colSpanCount
					} else {
						colCount += 1
					}
				}
			}
			return
		}
	})
	return
}

// 构建一个虚拟表格，用来表示那些格子被占用了
func BuildVirtualTable(rowCount, colCount int) (vTable [][]*model.TableCell) {
	vTable = make([][]*model.TableCell, 0)
	for i := 0; i < rowCount; i++ {
		rowCell := make([]*model.TableCell, colCount)
		vTable = append(vTable, rowCell)
	}
	return
}

// 根据html的表格，去构造一个被占用格子情况的表格
func SetUsedCellsInVTable(s *goquery.Selection, vTable [][]*model.TableCell) {
	s.Find("tbody tr").Each(func(rowIndex int, selection *goquery.Selection) {
		rowCellNodes := selection.Find("td").Nodes
		for _, cellNode := range rowCellNodes {
			rowSpan, colSpan := CalculateCellNodeSpan(cellNode.Attr)

			if rowSpan != 0 && colSpan == 0 { // 仅竖向合并的
				FillCellValue(rowIndex, cellNode.FirstChild.Data, true, rowSpan, 0, vTable)
			} else if colSpan != 0 && rowSpan == 0 { // 仅横向合并的
				for i := 0; i < colSpan; i++ {
					FillCellValue(rowIndex, cellNode.FirstChild.Data, false, 0, colSpan, vTable)
				}
			} else if colSpan != 0 && rowSpan != 0 { // 行列合并的
				for j := 0; j < colSpan; j++ {
					FillCellValue(rowIndex, cellNode.FirstChild.Data, true, rowSpan, colSpan, vTable)
				}
			} else { // 没有合并的
				FillCellValue(rowIndex, cellNode.FirstChild.Data, false, 0, 0, vTable)
			}
		}
	})
}

// 填充该行单元格
func FillCellValue(rowIndex int, value string, isVMerge bool, vMergeCount int, hMergeCount int, vTable [][]*model.TableCell) {
	if isVMerge {
		colIndex := -1
		for i := 0; i < vMergeCount; i++ {
			rowCell := vTable[rowIndex+i]
			if i == 0 {
				for key, cell := range rowCell {
					if cell == nil {
						c := &model.TableCell{
							RowIndex:      rowIndex,
							ColIndex:      key,
							Value:         value,
							IsVMerge:      true,
							HMerge:        hMergeCount,
							IsVMergeStart: true,
						}
						colIndex = key
						vTable[rowIndex][key] = c
						break
					}
				}
			} else {
				c := &model.TableCell{
					RowIndex: rowIndex,
					ColIndex: colIndex,
					Value:    value,
					IsVMerge: true,
					HMerge:   hMergeCount,
				}
				vTable[rowIndex+i][colIndex] = c
			}
		}
		return
	} else {
		rowCell := vTable[rowIndex]
		for key, cell := range rowCell {
			if cell == nil {
				c := &model.TableCell{
					RowIndex: rowIndex,
					ColIndex: key,
					HMerge:   hMergeCount,
					Value:    value,
				}
				vTable[rowIndex][key] = c
				return
			}
		}
	}
}

// 计算格子节点的行列合并数
func CalculateCellNodeSpan(attrs []html.Attribute) (rowSpan, colSpan int) {
	for _, attr := range attrs {
		if attr.Key == "colspan" {
			col, err := strconv.Atoi(attr.Val)
			if err != nil {
				log.Fatalln(err)
			}
			if col > 1 {
				colSpan = col
			}
		} else if attr.Key == "rowspan" {
			row, err := strconv.Atoi(attr.Val)
			if err != nil {
				log.Fatalln(err)
			}
			if row > 1 {
				rowSpan = row
			}
		}
	}
	return
}

// 判断节点是否有横向合并属性，如果有的话就返回合并数，如果没有的话就返回false
func IsCellHasColSpanAttr(node *html.Node) (has bool, colSpanCount int) {
	for _, attr := range node.Attr {
		if attr.Key == "colspan" {
			col, err := strconv.Atoi(attr.Val)
			if err != nil {
				log.Fatalln(err)
			}
			if col < 2 {
				colSpanCount += 1
			} else {
				colSpanCount += col
			}
			has = true
		}
	}
	return
}
