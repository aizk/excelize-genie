package excelizegenie

import (
	"bytes"
	"fmt"
	"io"

	"github.com/xuri/excelize/v2"
)

type Genie struct {
	File *excelize.File
}

type Options = excelize.Options

func OpenFile(filename string, opts ...Options) (g *Genie, err error) {
	file, err := excelize.OpenFile(filename, opts...)
	if err != nil {
		return
	}
	g = &Genie{
		File: file,
	}
	return
}

func OpenReader(r io.Reader, opts ...Options) (g *Genie, err error) {
	file, err := excelize.OpenReader(r, opts...)
	if err != nil {
		return
	}
	g = &Genie{
		File: file,
	}
	return
}

func NewGenie(file *excelize.File) *Genie {
	return &Genie{
		File: file,
	}
}

// CreateSheet Create New Sheet with datas
func (g *Genie) CreateSheet(sheetName string, datas [][]interface{}) (err error) {
	err = g.File.DeleteSheet(sheetName)
	if err != nil {
		return
	}

	_, err = g.File.NewSheet(sheetName)
	if err != nil {
		return
	}

	for i, data := range datas {
		err = g.File.SetSheetRow(sheetName, fmt.Sprintf("A%d", i+1), &data)
		if err != nil {
			return
		}
	}

	return
}

// Buffer Get Excel File Buffer
func (g *Genie) Buffer() (buf *bytes.Buffer, err error) {
	buf, err = g.File.WriteToBuffer()
	if err != nil {
		return
	}
	return
}

// GetIndexSheetRows Get Index Sheet Rows
func (g *Genie) GetIndexSheetRows(index int) (rows [][]string, err error) {
	return g.File.GetRows(g.File.WorkBook.Sheets.Sheet[index].Name)
}

// GetFirstSheetRows Get First Sheet All Rows
func (g *Genie) GetFirstSheetRows() (rows [][]string, err error) {
	return g.GetIndexSheetRows(1)
}
