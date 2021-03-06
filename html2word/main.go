package main

import (
	"fmt"
	"github.com/timliudream/officetools/html2word/model"
	"github.com/timliudream/officetools/html2word/wordstyle"
	"io/ioutil"
	"log"
	"os"
	"strings"

	"github.com/timliudream/officetools/html2word/utils"
	"golang.org/x/net/html"

	"github.com/PuerkitoBio/goquery"
)

func main() {
	sourcePath := "./html2word/htmltestset/普通测试文件.html"
	targetPath := "./html2word/test.docx"
	tmpHTMLPath := "./html2word/htmltmp/tmp.html"
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

	err = wordstyle.Doc.SaveToFile(targetPath)
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
				wordstyle.SetH(node.FirstChild.Data, tag)
			}
		} else if tag == "p" {
			if node.FirstChild != nil {
				if node.FirstChild.Type == html.TextNode {
					wordstyle.SetP(node.FirstChild.Data)
				} else if node.FirstChild.Type == html.ElementNode {
					pChild := node.FirstChild
					tag = pChild.DataAtom.String()
					if tag == "a" {
						if pChild.FirstChild != nil && pChild.FirstChild.Type == html.TextNode {
							wordstyle.SetHyperlink(pChild.FirstChild.Data)
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
							wordstyle.SetCode(node.FirstChild.Data)
						}
					}
				})
			}
		} else if tag == "table" {
			parseTable(s)
		} else if tag == "ul" { // 无序列表
			is := parseNotSortList(s)
			wordstyle.SetNotSortList(is, 0)
		} else if tag == "ol" { // 有序列表
			is := parseSortList(s)
			wordstyle.SetSortList(is, 0)
		}
	}
}

func parseImg(node *html.Node) {
	if node.FirstChild != nil {
		size := getImgSize(node)
		c := node.FirstChild.NextSibling.FirstChild
		attr := c.Attr[1]
		base64Str := strings.Replace(attr.Val, "\n", "", -1)
		base64Str, err := utils.StripMime(base64Str)
		if err != nil {
			log.Fatalln(err)
			return
		}
		imgPath := utils.Base2img(base64Str)
		err = wordstyle.SetImage(imgPath, size)
		if err != nil {
			log.Fatalln(err)
			return
		}
	}
}

func getImgSize(node *html.Node) string {
	size := ""
	for _, attr := range node.Attr {
		if attr.Key == "data-size" {
			size = attr.Val
		}
	}
	return size
}

func parseTable(s *goquery.Selection) {
	// 取行标题
	rowTitles := parseTableRowTitle(s)

	// 取列标题
	colTitles := parseTableColTitle(s)

	parseTableBody(s, rowTitles, colTitles)
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

func parseTableBody(s *goquery.Selection, rowTitles []*model.TableRowTitle, colTitles []*model.TableColTitle) {
	// 计算表格的行列数
	rowCount, colCount := utils.CalTableRowColCount(s)

	// 构造一个rowCount * colCount的矩阵，用来表示哪些格子被占用了
	vTable := utils.BuildVirtualTable(rowCount, colCount)
	fmt.Println(vTable)

	// 先测试一下看看格子占用情况是否正确
	utils.SetUsedCellsInVTable(s, vTable)

	if len(colTitles) > 0 {
		for rowIndex, row := range vTable {
			colTitleCells := make([]*model.TableCell, 0)
			cell := &model.TableCell{
				RowIndex:      rowIndex,
				ColIndex:      0,
				IsVMerge:      false,
				IsVMergeStart: false,
				HMerge:        0,
				Value:         colTitles[rowIndex].Title,
			}
			colTitleCells = append(colTitleCells, cell)
			row = append(colTitleCells, row...)
			vTable[rowIndex] = row
		}
	}
	if len(rowTitles) > 0 {
		t := make([][]*model.TableCell, 0)
		rowTitleCells := make([]*model.TableCell, 0)
		for colIndex, title := range rowTitles {
			cell := &model.TableCell{
				RowIndex:      0,
				ColIndex:      colIndex,
				IsVMerge:      false,
				IsVMergeStart: false,
				HMerge:        0,
				Value:         title.Title,
			}
			rowTitleCells = append(rowTitleCells, cell)
		}
		t = append(t, rowTitleCells)
		vTable = append(t, vTable...)
	}

	wordstyle.SetTable(vTable)
}

func parseNotSortList(s *goquery.Selection) []*model.NotSortItem {
	level := s.Children()
	levelList := make([]*model.NotSortItem, 0)
	for _, node := range level.Nodes {
		item := &model.NotSortItem{Value: strings.TrimRight(strings.Replace(node.FirstChild.Data, "\n", "", -1), " "), NotSortItemList: make([]*model.NotSortItem, 0)}
		levelList = append(levelList, item)
	}
	level.Each(func(i int, selection *goquery.Selection) {
		l := selection.Children()
		is := parseNotSortList(l)
		levelList[i].NotSortItemList = append(levelList[i].NotSortItemList, is...)
	})
	return levelList
}

func parseSortList(s *goquery.Selection) []*model.SortItem {
	level := s.Children()
	levelList := make([]*model.SortItem, 0)
	for _, node := range level.Nodes {
		item := &model.SortItem{Value: strings.TrimRight(strings.Replace(node.FirstChild.Data, "\n", "", -1), " "), SortItemList: make([]*model.SortItem, 0)}
		levelList = append(levelList, item)
	}
	level.Each(func(i int, selection *goquery.Selection) {
		l := selection.Children()
		is := parseSortList(l)
		levelList[i].SortItemList = append(levelList[i].SortItemList, is...)
	})
	return levelList
}
