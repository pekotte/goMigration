package excelhandler

import (
	"fmt"

	"github.com/xuri/excelize/v2"
	//"excelHandler"
)

func OpenAndShowCells(file string, sheet string) {
	f, err := excelize.OpenFile(file)
	if err != nil {
		fmt.Println(err)
		return
	}
	defer func() {
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()
	// Obtener valor de la celda por el nombre y el eje de la hoja de trabajo dado.
	/*cell, err := f.GetCellValue("Hoja1", "B2")
	if err != nil {
		fmt.Println(err)
		return
	}
	fmt.Println(cell)*/
	// Obtener todas las filas en el Sheet1.
	rows, err := f.GetRows(sheet)
	if err != nil {
		fmt.Println(err)
		return
	}
	for _, row := range rows {
		for _, colCell := range row {
			fmt.Print(colCell, "\t")
			fmt.Print("  ||  ")
		}
		fmt.Println()
		fmt.Println("======================================================================")
	}
}