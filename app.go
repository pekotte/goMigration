package main

import (
	//"fmt"

	//"github.com/xuri/excelize/v2"

	excelhandler "github.com/pekotte/goMigration/excelHandler"
)

func main() {

	excelhandler.OpenAndShowCells("MyTool.xlsx", "Hoja1")
}
