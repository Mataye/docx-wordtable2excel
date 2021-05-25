package main

import (
	"archive/zip"
	"encoding/xml"
	"io"
	"io/ioutil"
)

type Body struct {
	Paragraph []string `xml:"p>r>t"`
	Table     Table    `xml:"tbl"`
}

type Table struct {
	TableRow []TableRow `xml:"tr"`
}

type TableRow struct {
	TableColumn []TableColumn `xml:"tc"`
}

type TableColumn struct {
	Cell []string `xml:"p>r>t"`
}

type Document struct {
	XMLName xml.Name `xml:"document"`
	Body    Body     `xml:"body"`
}

func Extract(xmlContent string) (d Document) {
	err := xml.Unmarshal([]byte(xmlContent), &d)
	if err != nil {
		panic(err)
	}
	return
}

func UnpackDocx(filePath string) (*zip.ReadCloser, []*zip.File) {
	reader, err := zip.OpenReader(filePath)
	if err != nil {
		panic(err)
	}
	return reader, reader.File
}

func WordDocToString(reader io.Reader) (sContent string) {
	content, err := ioutil.ReadAll(reader)
	if err != nil {
		panic(err)
	}
	sContent = string(content)
	return
}

func RetrieveWordDoc(files []*zip.File) (file *zip.File) {
	for _, f := range files {
		if f.Name == "word/document.xml" {
			file = f
		}
	}
	return
}

func OpenWordDoc(doc zip.File) (rc io.ReadCloser) {
	rc, err := doc.Open()
	if err != nil {
		panic(err)
	}
	return
}
