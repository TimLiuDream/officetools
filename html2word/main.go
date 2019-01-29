package main

import (
	"io/ioutil"
	"log"
	"math"
	"os"
	"strconv"
	"strings"

	"github.com/timliudream/officetools/html2word/model"
	"github.com/timliudream/officetools/html2word/style"
	"github.com/timliudream/officetools/html2word/utils"
	"golang.org/x/net/html"

	"github.com/PuerkitoBio/goquery"
)

var cells []*model.TableCell

func main() {
	sourcePath := "test2.html"
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
			rowCount, colCount, cellMap := parseTable(s)
			err := style.SetTable(rowCount, colCount, cellMap, cells)
			if err != nil {
				log.Fatalln(err)
				return
			}
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

func parseTable(s *goquery.Selection) (rowCount, colCount int, tableCellMap map[string]*model.TableCell) {
	cells = make([]*model.TableCell, 0)
	tableCellMap = make(map[string]*model.TableCell)
	cellMap := make(map[string]string)
	tableRowSelection := s.Find("tbody tr")
	if tableRowSelection.Nodes != nil {
		rowCount = len(tableRowSelection.Nodes)
		colCount = 0
		tableRowSelection.Each(func(i int, selection *goquery.Selection) {
			cc := parseTableRow(i, selection, cellMap, tableCellMap)
			if cc > colCount {
				colCount = cc
			}
		})
	}
	return
}

func parseTableRow(rowIndex int, s *goquery.Selection, cellMap map[string]string, tableCellMap map[string]*model.TableCell) (colCount int) {
	tableColSeletion := s.Find("td")
	cellMergeCount := 0
	for colIndex, node := range tableColSeletion.Nodes {
		rowSpan := 0
		colSpan := 0
		for _, attr := range node.Attr {
			if attr.Key == "colspan" {
				col, err := strconv.Atoi(attr.Val)
				if err != nil {
					log.Fatalln(err)
				}
				if col == 1 {
					continue
				}
				colSpan = col
			} else if attr.Key == "rowspan" {
				row, err := strconv.Atoi(attr.Val)
				if err != nil {
					log.Fatalln(err)
				}
				if row == 1 {
					continue
				}
				rowSpan = row
			}
		}

		value := node.FirstChild.Data

		if rowSpan == 0 && colSpan == 0 {
			// 先要确定这个格子的索引
			for ci := 0; ci < math.MaxInt8; ci++ {
				cellKey := utils.GetCellKey(rowIndex, colIndex+ci)
				_, ok := cellMap[cellKey]
				if !ok {
					cellMap[cellKey] = value
					cell := &model.TableCell{RowIndex: rowIndex, ColIndex: colIndex + ci, Value: value}
					tableCellMap[cellKey] = cell
					cells = append(cells, cell)
					break
				}
			}
		}

		if rowSpan != 0 && colSpan == 0 {
			for ri := 0; ri < rowSpan; ri++ {
				cellKey := utils.GetCellKey(rowIndex+ri, colIndex+cellMergeCount)
				cellMap[cellKey] = value
				if rowIndex != rowIndex+rowSpan-1 {
					if !utils.IsCellInMergeCellScope(cellKey, tableCellMap) {
						cell := &model.TableCell{RowIndex: rowIndex + ri, ColIndex: colIndex + cellMergeCount, VMerge: rowSpan, Value: value}
						tableCellMap[cellKey] = cell
						cells = append(cells, cell)
					}
				}
			}
		} else if rowSpan == 0 && colSpan != 0 {
			for ci := 0; ci < colSpan; ci++ {
				cellKey := utils.GetCellKey(rowIndex, colIndex+ci+cellMergeCount)
				cellMap[cellKey] = value
				if colIndex != colSpan-1 {
					if !utils.IsCellInMergeCellScope(cellKey, tableCellMap) {
						cell := &model.TableCell{RowIndex: rowIndex, ColIndex: colIndex + ci + cellMergeCount, HMerge: colSpan, Value: value}
						tableCellMap[cellKey] = cell
						cells = append(cells, cell)
					}
				}
			}
			cellMergeCount += colSpan - 1
		} else if rowSpan != 0 && colSpan != 0 {
			// 计算每个格子的值
			for ri := 0; ri < rowSpan; ri++ {
				for ci := 0; ci < colSpan; ci++ {
					cellKey := utils.GetCellKey(rowIndex+ri, colIndex+ci+cellMergeCount)
					cellMap[cellKey] = value
					if !utils.IsCellInMergeCellScope(cellKey, tableCellMap) {
						cell := &model.TableCell{RowIndex: rowIndex + ri, ColIndex: colIndex + ci + cellMergeCount, VMerge: rowSpan, HMerge: colSpan, Value: value}
						tableCellMap[cellKey] = cell
						cells = append(cells, cell)
					}
				}
			}
			cellMergeCount += colSpan - 1
		}
	}
	colCount = cellMergeCount + len(tableColSeletion.Nodes)
	return
}
