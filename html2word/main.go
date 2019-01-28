package main

import (
	"fmt"
	"github.com/timliudream/officetools/html2word/logger"
	"github.com/timliudream/officetools/html2word/model"
	"github.com/timliudream/officetools/html2word/style"
	"github.com/timliudream/officetools/html2word/utils"
	"golang.org/x/net/html"
	"io/ioutil"
	"math"
	"os"
	"strconv"
	"strings"

	"github.com/PuerkitoBio/goquery"
)

func main() {
	sourcePath := "test2.html"
	targetPath := "test.docx"
	tmpHtmlPath := "htmltmp/tmp.html"
	file, err := os.Open(sourcePath)
	if err != nil {
		logger.Error.Println(err)
		return
	}
	htmlDoc, err := goquery.NewDocumentFromReader(file)
	if err != nil {
		logger.Error.Println(err)
		return
	}

	// 先对文档做markdown和code处理
	htmlDoc.Find("div[class=ones-marked-card]").Each(func(i int, selection *goquery.Selection) {
		err, output := utils.ConvertMarkdownToHTML(selection.Text())
		if err != nil {
			logger.Error.Println(err)
			return
		}
		// 不知道为什么不做截取操作的话，是取不到body的内容的
		outputs := strings.Split(output, "body")
		realOutput := strings.TrimLeft(outputs[1], ">")
		realOutput = strings.TrimRight(realOutput, "</")
		fmt.Println(realOutput)
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
		logger.Error.Println(err)
		return
	}
	content = html.UnescapeString(content)

	err = ioutil.WriteFile(tmpHtmlPath, []byte(content), 0644)
	if err != nil {
		logger.Error.Println(err)
		return
	}

	// 正式处理
	file, err = os.Open(tmpHtmlPath)
	if err != nil {
		logger.Error.Println(err)
		return
	}
	htmlDoc, err = goquery.NewDocumentFromReader(file)
	if err != nil {
		logger.Error.Println(err)
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
		logger.Error.Println(err)
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
			rowCount, colCount, mergeCellScopeMap := parseTable(node, s)
			err := style.SetTable(rowCount, colCount, mergeCellScopeMap)
			if err != nil {
				logger.Error.Println(err)
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
			logger.Error.Println(err)
			return
		}
		imgPath := utils.Base2img(base64Str)
		err = style.SetImage(imgPath)
		if err != nil {
			logger.Error.Println(err)
			return
		}
	}
}

func parseTable(node *html.Node, s *goquery.Selection) (rowCount, colCount int, mergeCellScopeMap map[string]*model.MergeCellScope) {
	mergeCellScopeMap = make(map[string]*model.MergeCellScope)

	tableRowSelection := s.Find("tbody tr")
	if tableRowSelection.Nodes != nil {
		rowCount = len(tableRowSelection.Nodes)
		colCount = 0
		tableRowSelection.Each(func(i int, selection *goquery.Selection) {
			cc := parseTableRow(i, selection, mergeCellScopeMap)
			if cc > colCount {
				colCount = cc
			}
		})
	}
	return
}

func parseTableRow(rowIndex int, s *goquery.Selection, mergeCellScopeMap map[string]*model.MergeCellScope) (colCount int) {
	tableColSeletion := s.Find("td")
	cellMergeCount := 0
	for colIndex, node := range tableColSeletion.Nodes {
		rowSpan := 0
		colSpan := 0
		for _, attr := range node.Attr {
			if attr.Key == "colspan" {
				col, err := strconv.Atoi(attr.Val)
				if err != nil {
					logger.Error.Println(err)
				}
				colSpan = col
			} else if attr.Key == "rowspan" {
				row, err := strconv.Atoi(attr.Val)
				if err != nil {
					logger.Error.Println(err)
				}
				rowSpan = row
			}
		}

		if rowSpan == 0 && colSpan == 0 {
			for ci := 0; ci < math.MaxInt8; ci++ {
				cellKey := utils.GetCellKey(rowIndex, colIndex+ci)
				_, ok := mergeCellScopeMap[cellKey]
				if !ok {
					mergeCellScopeMap[cellKey] = &model.MergeCellScope{Value: node.FirstChild.Data}
					break
				}
			}
		}

		if rowSpan != 0 && colSpan == 0 {
			for ri := 0; ri < rowSpan; ri++ {
				cellKey := utils.GetCellKey(rowIndex+ri, colIndex+cellMergeCount)
				if rowIndex != rowIndex+rowSpan-1 {
					if !utils.IsCellInMergeCellScope(cellKey, mergeCellScopeMap) {
						mergeCellScopeMap[cellKey] = &model.MergeCellScope{RowScope: model.RowScope{Start: rowIndex, End: rowIndex + rowSpan - 1}, Value: node.FirstChild.Data}
					}
				}
			}
		} else if rowSpan == 0 && colSpan != 0 {
			for ci := 0; ci < colSpan; ci++ {
				cellKey := utils.GetCellKey(rowIndex, colIndex+ci+cellMergeCount)
				if colIndex != colSpan-1 {
					if !utils.IsCellInMergeCellScope(cellKey, mergeCellScopeMap) {
						mergeCellScopeMap[cellKey] = &model.MergeCellScope{ColScope: model.ColScope{Start: colIndex, End: colIndex + colSpan - 1}, Value: node.FirstChild.Data}
					}
				}
			}
			cellMergeCount += colSpan - 1
		} else if rowSpan != 0 && colSpan != 0 {
			// 计算每个格子的值
			for ri := 0; ri < rowSpan; ri++ {
				for ci := 0; ci < colSpan; ci++ {
					cellKey := utils.GetCellKey(rowIndex+ri, colIndex+ci+cellMergeCount)
					if !utils.IsCellInMergeCellScope(cellKey, mergeCellScopeMap) {
						var rs model.RowScope
						var cs model.ColScope
						if rowIndex != rowIndex+rowSpan-1 {
							rs = model.RowScope{Start: rowIndex, End: rowIndex + rowSpan - 1}
						}
						if colIndex != colSpan-1 {
							cs = model.ColScope{Start: colIndex, End: colIndex + colSpan - 1}
						}
						mergeCellScopeMap[cellKey] = &model.MergeCellScope{RowScope: rs, ColScope: cs, Value: node.FirstChild.Data}
					}
				}
			}
			cellMergeCount += colSpan - 1
		}
	}
	colCount = cellMergeCount + len(tableColSeletion.Nodes)
	return
}
