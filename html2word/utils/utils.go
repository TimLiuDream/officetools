package utils

import (
	"bytes"
	"encoding/base64"
	"errors"
	"fmt"
	"github.com/PuerkitoBio/goquery"
	"github.com/russross/blackfriday"
	"github.com/satori/go.uuid"
	"io/ioutil"
	"log"
	"regexp"
	"strings"
)

// base64字符串转图片并保存
func Base2img(base64Str string) (imgPath string) {
	UUID, err := uuid.NewV4()
	if err != nil {
		log.Fatalln(err)
	}
	imgPath = fmt.Sprintf("./html2word/image/%s", UUID.String()+".jpg")
	ddd, _ := base64.RawStdEncoding.DecodeString(base64Str)
	err = ioutil.WriteFile(imgPath, ddd, 0666)
	if err != nil {
		log.Fatalln(err)
	}
	return
}

func StripMime(combined string) (string, error) {
	re := regexp.MustCompile("data:(.*);base64,(.*)")
	parts := re.FindStringSubmatch(combined)

	if len(parts) < 3 {
		return "", errors.New("invalid base64 input")
	}

	data := parts[2]
	return data, nil
}

func ConvertMarkdownToHTML(input string) (error, string) {
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
		return err, ""
	}
	doc.Find("body").Each(func(i int, selection *goquery.Selection) {
		ret, _ := selection.Html()
		ret = strings.Replace(ret, "<pre>", "<blockquote><pre>", -1)
		ret = strings.Replace(ret, "</pre>", "</blockquote></pre>", -1)
		selection.SetHtml(ret)
	})
	html, _ := doc.Html()
	return nil, html
}
