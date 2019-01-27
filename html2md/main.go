package main

import (
	"fmt"
	"io/ioutil"
	"log"

	"github.com/lunny/html2md"
)

func main() {
	path := "/Users/tim/Documents/test.html"
	mdPath := "/Users/tim/Documents/test.md"
	bytes, err := ioutil.ReadFile(path)
	if err != nil {
		log.Fatal(err)
	}

	md := html2md.Convert(string(bytes))

	fmt.Println(md)

	err = ioutil.WriteFile(mdPath, []byte(md), 0755)
	if err != nil {
		log.Fatal(err)
	}
}
