package excelizegenie

import (
	"github.com/xuri/excelize/v2"
)

// CreateExcelSimple Fast Create an Excel With One Sheet
func CreateExcelSimple(sheetName string, datas [][]interface{}) (g *Genie, err error) {
	f := excelize.NewFile()
	defer func() {
		if err = f.Close(); err != nil {
			return
		}
	}()

	g = NewGenie(f)
	err = g.CreateSheet(sheetName, datas)
	if err != nil {
		return
	}
	return
}
