package utils

import (
	"bytes"
	"encoding/base64"
	"errors"
	"fmt"
	"io/ioutil"
	"log"
	"regexp"
	"strconv"
	"strings"

	"github.com/PuerkitoBio/goquery"
	"github.com/russross/blackfriday"
	"github.com/satori/go.uuid"
)

// Base2img base64字符串转图片并保存
func Base2img(base64Str string) (imgPath string) {
	UUID, err := uuid.NewV4()
	if err != nil {
		log.Fatalln(err)
		return
	}
	imgPath = fmt.Sprintf("./html2word/image/%s", UUID.String()+".jpg")
	ddd, _ := base64.RawStdEncoding.DecodeString(base64Str)
	err = ioutil.WriteFile(imgPath, ddd, 0666)
	if err != nil {
		log.Fatalln(err)
		return
	}
	return
}

// StripMime 取出base64字符串中图片有效字符串
func StripMime(combined string) (string, error) {
	re := regexp.MustCompile("data:(.*);base64,(.*)")
	parts := re.FindStringSubmatch(combined)

	if len(parts) < 3 {
		return "", errors.New("invalid base64 input")
	}

	data := parts[2]
	return data, nil
}

// ConvertMarkdownToHTML 将markdown转换成html
func ConvertMarkdownToHTML(input string) (string, error) {
	var renderer blackfriday.Renderer
	extensions := 0
	extensions |= blackfriday.EXTENSION_NO_INTRA_EMPHASIS
	extensions |= blackfriday.EXTENSION_TABLES
	extensions |= blackfriday.EXTENSION_FENCED_CODE
	extensions |= blackfriday.EXTENSION_AUTOLINK
	extensions |= blackfriday.EXTENSION_STRIKETHROUGH
	extensions |= blackfriday.EXTENSION_SPACE_HEADERS
	extensions |= blackfriday.EXTENSION_BACKSLASH_LINE_BREAK

	htmlFlags := 0
	htmlFlags |= blackfriday.HTML_COMPLETE_PAGE

	renderer = blackfriday.HtmlRenderer(htmlFlags, "", "")
	output := blackfriday.Markdown([]byte(input), renderer, extensions)
	doc, err := goquery.NewDocumentFromReader(bytes.NewReader(output))
	if err != nil {
		return "", err
	}
	doc.Find("body").Each(func(i int, selection *goquery.Selection) {
		ret, _ := selection.Html()
		ret = strings.Replace(ret, "<pre>", "<blockquote><pre>", -1)
		ret = strings.Replace(ret, "</pre>", "</blockquote></pre>", -1)
		selection.SetHtml(ret)
	})
	html, _ := doc.Html()
	return html, nil
}

// GetCellKey 根据行列索引给出对应的cellMap的key
func GetCellKey(rowIndex, colIndex int) string {
	return strconv.Itoa(rowIndex) + "," + strconv.Itoa(colIndex)
}

// GetRowColByCellKey 将cellMap的key分解成行列索引
func GetRowColByCellKey(cellKey string) (row, col int) {
	rowColCouple := strings.Split(cellKey, ",")
	rowStr := rowColCouple[0]
	colStr := rowColCouple[1]
	row, err := strconv.Atoi(rowStr)
	if err != nil {
		log.Fatalln(err)
		return
	}
	col, err = strconv.Atoi(colStr)
	if err != nil {
		log.Fatalln(err)
		return
	}
	return
}

// IsCellInMergeCellScope 判断单元格是不是在合并的map中已经包含了
//func IsCellInMergeCellScope(cellKey string, tableCellMap map[string]*model.TableCell) (result bool) {
//	for key, value := range tableCellMap {
//		row, col := GetRowColByCellKey(key)
//		rowStart := row
//		rowEnd := row + value.VMerge
//		colStart := col
//		colEnd := col + value.HMerge
//
//		cellRow, cellCol := GetRowColByCellKey(cellKey)
//		if cellRow >= rowStart && cellRow <= rowEnd && cellCol >= colStart && cellCol <= colEnd {
//			result = true
//			return
//		}
//	}
//	return
//}
