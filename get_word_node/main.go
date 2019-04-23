package main

import (
	"log"

	"baliance.com/gooxml/document"
)

func main() {
	path := "/Users/tim/go/src/github.com/timliudream/officetools/get_word_node/背景色.docx"

	doc, err := document.Open(path)
	if err != nil {
		log.Fatalf("error opening document: %s", err)
	}

	paragraphs := []document.Paragraph{}
	runs := []document.Run{}
	for _, p := range doc.Paragraphs() {
		for _, r := range p.Runs() {
			paragraphs = append(paragraphs, p)
			runs = append(runs, r)
		}
	}
}
