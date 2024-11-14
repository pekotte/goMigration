package main

import (
	"fmt"

	excelhandler "github.com/pekotte/goMigration/excelHandler"
)

func main() {

	//excelhandler.OpenAndShowAllCells("MyTool.xlsx", "Hoja1")

	type sEmployee struct {
		ui32idAgresso uint32
		strName       string
		strDomain     string
		strCmpyStDate string
		strEmail      string
		bPcAlten      bool
		iOficina      int
		strLocation   string
		strLocAssoc   string
		ui32Salary    uint32
		ui32Bond      uint32
		bRestaurant   bool
		bInsurance    bool
		ui8Holidays   uint8
		strComments   string
		ui8SalaryDay  uint8
		strBmADC      string
		strCategory   string
		strN1         string
		strN2         string
		strIntBm      string
	}

	type sProyect struct {
		strCode      string
		strName      string
		strStatus    string
		strStartDate string
	}

	type sClient struct {
		strCode string
		strName string
	}

	var newClient sClient = sClient{
		excelhandler.OpenAndShowOneCells("MyTool.xlsx", "Hoja1", "D2"),
		excelhandler.OpenAndShowOneCells("MyTool.xlsx", "Hoja1", "E2")}

	fmt.Println(newClient)

}
