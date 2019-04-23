package main

import (
	"fmt"
	"log"
	"time"

	"baliance.com/gooxml/document"
)

func main() {
	doc, err := document.Open("/Users/tim/go/src/github.com/timliudream/officetools/doc_properties/背景色.docx")
	if err != nil {
		log.Fatalf("error opening document: %s", err)
	}

	cp := doc.CoreProperties
	// You can read properties from the document
	fmt.Println("Title:", cp.Title())
	fmt.Println("Author:", cp.Author())
	fmt.Println("Description:", cp.Description())
	fmt.Println("Last Modified By:", cp.LastModifiedBy())
	fmt.Println("Category:", cp.Category())
	fmt.Println("Content Status:", cp.ContentStatus())
	fmt.Println("Created:", cp.Created())
	fmt.Println("Modified:", cp.Modified())

	// And change them as well
	cp.SetTitle("CP Invoices")
	cp.SetAuthor("John Doe")
	cp.SetCategory("Invoices")
	cp.SetContentStatus("Draft")
	cp.SetLastModifiedBy("Jane Smith")
	cp.SetCreated(time.Now())
	cp.SetModified(time.Now())
	doc.SaveToFile("./doc_properties/document.docx")
}
