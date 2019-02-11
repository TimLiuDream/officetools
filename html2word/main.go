package main

import (
	"fmt"
	"github.com/timliudream/officetools/html2word/model"
	"io/ioutil"
	"log"
	"os"
	"strconv"
	"strings"

	"github.com/timliudream/officetools/html2word/style"
	"github.com/timliudream/officetools/html2word/utils"
	"golang.org/x/net/html"

	"github.com/PuerkitoBio/goquery"
)

func main() {
	sourcePath := "htmltestset/多种合并方式的表格2.html"
	targetPath := "test.docx"
	tmpHTMLPath := "htmltmp/tmp.html"
	file, err := os.Open(sourcePath)
	if err != nil {
		log.Fatalln(err)
		return
	}
	htmlDoc, err := goquery.NewDocumentFromReader(file)
	if err != nil {
		log.Fatalln(err)
		return
	}

	// 先对文档做markdown和code处理
	htmlDoc.Find("div[class=ones-marked-card]").Each(func(i int, selection *goquery.Selection) {
		output, err := utils.ConvertMarkdownToHTML(selection.Text())
		if err != nil {
			log.Fatalln(err)
			return
		}
		// 不知道为什么不做截取操作的话，是取不到body的内容的
		outputs := strings.Split(output, "body")
		realOutput := strings.TrimLeft(outputs[1], ">")
		realOutput = strings.TrimRight(realOutput, "</")
		selection.SetText(realOutput)
	})
	htmlDoc.Find("div[class=ones-code-card]").Each(func(i int, selection *goquery.Selection) {
		ret, _ := selection.Html()
		ret = strings.Replace(ret, "<pre>", "<blockquote><pre>", -1)
		ret = strings.Replace(ret, "</pre>", "</blockquote></pre>", -1)
		selection.SetHtml(ret)
	})
	content, err := htmlDoc.Html()
	if err != nil {
		log.Fatalln(err)
		return
	}
	content = html.UnescapeString(content)

	err = ioutil.WriteFile(tmpHTMLPath, []byte(content), 0644)
	if err != nil {
		log.Fatalln(err)
		return
	}

	// 正式处理
	file, err = os.Open(tmpHTMLPath)
	if err != nil {
		log.Fatalln(err)
		return
	}
	htmlDoc, err = goquery.NewDocumentFromReader(file)
	if err != nil {
		log.Fatalln(err)
		return
	}

	rootChildren := htmlDoc.Find("body").Children()
	rootChildren.Each(func(i int, selection *goquery.Selection) {
		for _, node := range selection.Nodes {
			parseElement(node, selection)
		}
	})

	err = style.Doc.SaveToFile(targetPath)
	if err != nil {
		log.Fatalln(err)
		return
	}
}

func parseElement(node *html.Node, s *goquery.Selection) {
	if node.Type == html.ElementNode {
		tag := node.DataAtom.String()
		if strings.HasPrefix(tag, "h") {
			if node.FirstChild != nil && node.FirstChild.Type == html.TextNode {
				style.SetH(node.FirstChild.Data, tag)
			}
		} else if tag == "p" {
			if node.FirstChild != nil {
				if node.FirstChild.Type == html.TextNode {
					style.SetP(node.FirstChild.Data)
				} else if node.FirstChild.Type == html.ElementNode {
					pChild := node.FirstChild
					tag = pChild.DataAtom.String()
					if tag == "a" {
						if pChild.FirstChild != nil && pChild.FirstChild.Type == html.TextNode {
							style.SetHyperlink(pChild.FirstChild.Data)
						}
					}
				}
			}
		} else if tag == "figure" {
			parseImg(node)
		} else if tag == "div" {
			if node.Attr[0].Val == "ones-marked-card" {
				// markdown
				if node.FirstChild != nil && node.FirstChild.Type == html.TextNode {
					n := node.FirstChild.NextSibling
					parseElement(n, s)
				}
			} else if node.Attr[0].Val == "ones-code-card" {
				// code
				s.Find("div code").Each(func(i int, selection *goquery.Selection) {
					for _, node := range selection.Nodes {
						if node.FirstChild != nil && node.FirstChild.Type == html.TextNode {
							style.SetCode(node.FirstChild.Data)
						}
					}
				})
			}
		} else if tag == "table" {
			parseTable(s)
			//err := style.SetTable(rowCount, colCount, cellMap, cells)
			//if err != nil {
			//	log.Fatalln(err)
			//	return
			//}
		}
	}
}

func parseImg(node *html.Node) {
	if node.FirstChild != nil {
		c := node.FirstChild.NextSibling.FirstChild
		attr := c.Attr[1]
		base64Str := strings.Replace(attr.Val, "\n", "", -1)
		base64Str, err := utils.StripMime(base64Str)
		if err != nil {
			log.Fatalln(err)
			return
		}
		imgPath := utils.Base2img(base64Str)
		err = style.SetImage(imgPath)
		if err != nil {
			log.Fatalln(err)
			return
		}
	}
}

func parseTable(s *goquery.Selection) {
	// 取行标题
	rowTitles := parseTableRowTitle(s)

	// 取列标题
	colTitles := parseTableColTitle(s)
	fmt.Println(rowTitles)
	fmt.Println(colTitles)

	parseTableBody(s)
	return
}

func parseTableRowTitle(s *goquery.Selection) (rowTitles []*model.TableRowTitle) {
	rowTitles = make([]*model.TableRowTitle, 0)
	rowTitleNodes := s.Find("thead tr th").Nodes
	if len(rowTitleNodes) != 0 {
		for index, node := range rowTitleNodes {
			if node.FirstChild != nil {
				rowTitle := &model.TableRowTitle{ColIndex: index, Title: node.FirstChild.Data}
				rowTitles = append(rowTitles, rowTitle)
			} else {
				rowTitle := &model.TableRowTitle{ColIndex: index, Title: ""}
				rowTitles = append(rowTitles, rowTitle)
			}
		}
	}
	return
}

func parseTableColTitle(s *goquery.Selection) (colTitles []*model.TableColTitle) {
	colTitles = make([]*model.TableColTitle, 0)
	colTitleNodes := s.Find("tbody tr th").Nodes
	if len(colTitleNodes) != 0 {
		for index, node := range colTitleNodes {
			if node.FirstChild != nil {
				colTitle := &model.TableColTitle{RowIndex: index, Title: node.FirstChild.Data}
				colTitles = append(colTitles, colTitle)
			} else {
				colTitle := &model.TableColTitle{RowIndex: index, Title: ""}
				colTitles = append(colTitles, colTitle)
			}
		}
	}
	return
}

func parseTableBody(s *goquery.Selection) (tableCells [][]*model.TableCell) {
	// 计算表格的行列数
	rowCount, colCount := calTableRowColCount(s)

	// 构造一个rowCount * colCount的矩阵，用来表示哪些格子被占用了
	vTable := buildVirtualTable(rowCount, colCount)
	fmt.Println(vTable)

	tableCells = make([][]*model.TableCell, 0)

	// 先测试一下看看格子占用情况是否正确
	setUsedCellsInVTable(s, vTable)

	// 先遍历行
	rows := s.Find("tbody tr")
	rows.Each(func(rowIndex int, selection *goquery.Selection) {
		// 遍历行中的列
		//rowCells := parseTableRow(rowIndex, selection, rowCellMap)
		//tableCells = append(tableCells, rowCells)
	})
	return
}

// 计算表格的行列数
func calTableRowColCount(s *goquery.Selection) (rowCount, colCount int) {
	rowSelection := s.Find("tbody tr")
	rowCount = len(rowSelection.Nodes)
	rowSelection.Each(func(i int, selection *goquery.Selection) {
		if i == 0 {
			cellSelection := selection.Find("td")
			for _, node := range cellSelection.Nodes {
				if node.Attr == nil {
					colCount += 1
				} else {
					has, colSpanCount := isCellHasColSpanAttr(node)
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
func buildVirtualTable(rowCount, colCount int) (vTable [][]*model.TableCell) {
	vTable = make([][]*model.TableCell, 0)
	for i := 0; i < rowCount; i++ {
		rowCell := make([]*model.TableCell, colCount)
		vTable = append(vTable, rowCell)
	}
	return
}

// 根据html的表格，去构造一个被占用格子情况的表格
func setUsedCellsInVTable(s *goquery.Selection, vTable [][]*model.TableCell) {
	s.Find("tbody tr").Each(func(rowIndex int, selection *goquery.Selection) {
		rowCellNodes := selection.Find("td").Nodes
		for _, cellNode := range rowCellNodes {
			rowSpan, colSpan := calculateCellNodeSpan(cellNode.Attr)

			if rowSpan != 0 && colSpan == 0 { // 仅竖向合并的
				fillCellValue(rowIndex, cellNode.FirstChild.Data, true, rowSpan, vTable)
			} else if colSpan != 0 && rowSpan == 0 { // 仅横向合并的
				for i := 0; i < colSpan; i++ {
					fillCellValue(rowIndex, cellNode.FirstChild.Data, false, 0, vTable)
				}
			} else if colSpan != 0 && rowSpan != 0 { // 行列合并的
				for j := 0; j < colSpan; j++ {
					fillCellValue(rowIndex, cellNode.FirstChild.Data, true, rowSpan, vTable)
				}
			} else { // 没有合并的
				fillCellValue(rowIndex, cellNode.FirstChild.Data, false, 0, vTable)
			}
		}
	})
}

// 填充该行单元格
func fillCellValue(rowIndex int, value string, isVMerge bool, vMergeCount int, vTable [][]*model.TableCell) {
	if isVMerge {
		colIndex := -1
		for i := 0; i < vMergeCount; i++ {
			rowCell := vTable[rowIndex+i]
			if i == 0 {
				for key, cell := range rowCell {
					if cell == nil {
						c := &model.TableCell{
							RowIndex: rowIndex,
							ColIndex: key,
							Value:    value,
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
					Value:    value,
				}
				vTable[rowIndex][key] = c
				return
			}
		}
	}
}

// 计算格子节点的行列合并数
func calculateCellNodeSpan(attrs []html.Attribute) (rowSpan, colSpan int) {
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
func isCellHasColSpanAttr(node *html.Node) (has bool, colSpanCount int) {
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
