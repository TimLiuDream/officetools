package main

import (
	"io/ioutil"
	"log"
	"os/exec"
)

func main() {
	originPath := "/Users/tim/Documents/Q94F7EBW_5jQqyKd3_1547603150377079.html"
	targetPath := "/Users/tim/Documents/Q94F7EBW_5jQqyKd3_1547603150377079.docx"
	referenceDocPath := "/Users/tim/Documents/custom-reference.docx"
	cmd := exec.Command("pandoc", "-f", "HTML", originPath, "-o", targetPath, "--reference-doc", referenceDocPath)
	err := cmd.Run()
	if err != nil {
		log.Fatal(err)
	}
	_, err = ioutil.ReadFile(targetPath)
	if err != nil {
		log.Fatal(err)
	}
}
