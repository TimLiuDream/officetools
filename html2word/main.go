package main

import (
	"fmt"
	"github.com/timliudream/GolangTraining/html2word/style"
	"github.com/timliudream/GolangTraining/html2word/utils"
	"golang.org/x/net/html"
	"io/ioutil"
	"log"
	"os"
	"strings"

	"github.com/PuerkitoBio/goquery"
)

func main() {
	sourcePath := "./html2word/test2.html"
	targetPath := "./html2word/test.docx"
	tmpHtmlPath := "./html2word/htmltmp/tmp.html"
	file, err := os.Open(sourcePath)
	if err != nil {
		log.Fatalln(err)
	}
	htmlDoc, err := goquery.NewDocumentFromReader(file)
	if err != nil {
		log.Fatalln(err)
	}

	// 先对文档做markdown和code处理
	htmlDoc.Find("div[class=ones-marked-card]").Each(func(i int, selection *goquery.Selection) {
		err, output := utils.ConvertMarkdownToHTML(selection.Text())
		if err != nil {
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
		return
	}
	content = html.UnescapeString(content)

	err = ioutil.WriteFile(tmpHtmlPath, []byte(content), 0644)
	if err != nil {
		log.Fatalln(err)
	}

	// 正式处理
	file, err = os.Open(tmpHtmlPath)
	if err != nil {
		log.Fatalln(err)
	}
	htmlDoc, err = goquery.NewDocumentFromReader(file)
	if err != nil {
		log.Fatalln(err)
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
		}
		imgPath := utils.Base2img(base64Str)
		err = style.SetImage(imgPath)
		if err != nil {
			log.Fatalln(err)
		}
	}
}
