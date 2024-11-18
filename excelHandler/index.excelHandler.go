package excelhandler

import (
	"fmt"
	"strconv"

	"github.com/xuri/excelize/v2"
)

var myFile string = "MyTool.xlsx"
var mySheet string = "Hoja1"

const (
	CLIENT_COD     string = "D"
	CLIENT_NAME    string = "E"
	PROYECT_COD    string = "H"
	PROYECT_NAME   string = "G"
	PROYECT_STATUS string = "J"
	PROYECT_SDATE  string = "L"

	EMPLOYEE_AGRESSO    string = "A"
	EMPLOYEE_NAME       string = "B"
	EMPLOYEE_DOMAIN     string = "C"
	EMPLOYEE_CMPYSTDATE string = "K"
	EMPLOYEE_EMAIL      string = "M"
	EMPLOYEE_PC_ALTEN   string = "N"
	EMPLOYEE_OFFICE     string = "O"
	EMPLOYEE_LOCATION   string = "P"
	EMPLOYEE_LOC_ASSOC  string = "Q"
	EMPLOYEE_SALARY     string = "R"
	EMPLOYEE_BOND       string = "S"
	EMPLOYEE_RESTAURANT string = "T"
	EMPLOYEE_INSURANCE  string = "U"
	EMPLOYEE_HOLIDAYS   string = "V"
	EMPLOYEE_COMMENTS   string = "W"
	EMPLOYEE_SALARY_DAY string = "X"
	EMPLOYEE_BM_ADC     string = "Y"
	EMPLOYEE_CATEGORY   string = "Z"
	EMPLOYEE_N1         string = "AA"
	EMPLOYEE_N2         string = "AB"
	EMPLOYEE_INT_BM     string = "AC"
)

func ShowAllCells() {
	f, err := excelize.OpenFile(myFile)
	if err != nil {
		fmt.Println("(1) ", err)
		return
	}
	defer func() {
		if err := f.Close(); err != nil {
			fmt.Println("(2) ", err)
			return
		}
	}()

	rows, err := f.GetRows(mySheet)
	if err != nil {
		fmt.Println("(3) ", err)
		return
	}
	for _, row := range rows {
		for _, colCell := range row {
			fmt.Print(colCell, "\t")
			fmt.Print("  ||  ")
		}
		fmt.Println()
	}
}

func GetFileRows() *excelize.Rows {
	f, err := excelize.OpenFile(myFile)
	if err != nil {
		fmt.Println("(1) ", err)
		return nil
	}
	defer func() {
		if err := f.Close(); err != nil {
			fmt.Println("(2) ", err)
			return
		}
	}()

	rows, err := f.Rows(mySheet)
	if err != nil {
		fmt.Println(err)
		return nil
	}

	return rows
}

func OpenAndShowOneCells(cell string) string {
	// Obtener valor de la celda por el nombre y el eje de la hoja de trabajo dado.
	f, err := excelize.OpenFile(myFile)
	if err != nil {
		fmt.Println("(1) ", err)
		return ""
	}
	defer func() {
		if err := f.Close(); err != nil {
			fmt.Println("(2) ", err)
			return
		}
	}()

	mycell, err := f.GetCellValue(mySheet, cell)
	if err != nil {
		fmt.Println(err)
		return ""
	}
	return mycell
}

func GetOneData(dataType string, row int) string {
	return OpenAndShowOneCells(dataType + strconv.Itoa(row))
}
