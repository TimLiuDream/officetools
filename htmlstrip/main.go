package main

import (
	"bytes"
	"fmt"
	"strings"

	"github.com/PuerkitoBio/goquery"
	"golang.org/x/net/html"
)

func html2text(s *goquery.Document) string {
	var buf bytes.Buffer
	firstLine := true
	// Slightly optimized vs calling Each: no single selection object created
	var f func(*html.Node)
	f = func(n *html.Node) {
		if n.Type == html.TextNode {
			// Keep newlines and spaces, like jQuery
			//fmt.Printf("data: [%s]\n", n.Data)
			// 过滤全为空格的行
			if strings.TrimSpace(n.Data) != "" {
				buf.WriteString(n.Data)
			}
		} else if n.Type == html.ElementNode {
			tag := n.DataAtom.String()
			if tag == "p" || tag == "li" {
				if firstLine {
					firstLine = false
				} else {
					buf.WriteString("\n")
				}
			} else if tag == "figure" {
				buf.WriteString("[图片]")
				return
			}
		}
		if n.FirstChild != nil {
			for c := n.FirstChild; c != nil; c = c.NextSibling {
				f(c)
			}
		}
	}
	for _, n := range s.Nodes {
		f(n)
	}
	return buf.String()
}

func HtmlStrip(html string) (string, error) {
	buf := bytes.NewBufferString(html)
	doc, err := goquery.NewDocumentFromReader(buf)
	if err != nil {
		return "", err
	}
	if doc == nil {
		return "", nil
	}
	return html2text(doc), nil
}

func main() {
	html := "<p>操作步骤：</p>"
	html += "<p>adfa<span style=\"color:#808001\">sfasdfasdfasdfa</span>sdfasdfas</p>"
	html += "<p>&nbsp;</p>"
	html += "<p>预期结果：</p>"
	html += "<p>&nbsp;</p>"
	html += "<p>222</p>"
	s, err := HtmlStrip(html)
	if err != nil {
		fmt.Println(err)
		return
	}
	fmt.Println(s)
}
