package main

import (
	"fmt"

	excelhandler "github.com/pekotte/goMigration/excelHandler"
)

func main() {

	//excelhandler.OpenAndShowAllCells("MyTool.xlsx", "Hoja1")

	type sEmployee struct {
		strIdAgresso  string
		strName       string
		strDomain     string
		strCmpyStDate string
		strEmail      string
		strPcAlten    string
		strOficina    string
		strLocation   string
		strLocAssoc   string
		strSalary     string
		strBond       string
		strRestaurant string
		strInsurance  string
		strHolidays   string
		strComments   string
		ui8SalaryDay  string
		strBmADC      string
		strCategory   string
		strN1         string
		strN2         string
		strIntBm      string
	}

	type sProyectMain struct {
		strName string
	}

	type sProyect struct {
		strCode      string
		strName      string
		strStatus    string
		strStartDate string
	}

	type sClient struct {
		strName string
	}

	type sClientInt struct {
		strCode string
		strName string
	}

	type sCompleteData struct {
		clientInt   sClientInt
		client      sClient
		proyectMain sProyectMain
		proyect     sProyect
		employee    sEmployee
	}

	rows := excelhandler.GetFileRows()

	sliceExcelData := make([]sCompleteData, 3, 500)

	i := 2
	for rows.Next() {
		//fmt.Println("------------------------------------------------")
		var newClientInt sClientInt = sClientInt{
			excelhandler.GetOneData(excelhandler.CLIENT_COD_INT, i),
			excelhandler.GetOneData(excelhandler.CLIENT_NAME_INT, i)}

		if newClientInt.strName == "" {
			continue
		}

		//fmt.Println(newClientInt)

		var newClient sClient = sClient{
			excelhandler.GetOneData(excelhandler.CLIENT_NAME, i)}
		//fmt.Println(newClient)

		var newProyectMain sProyectMain = sProyectMain{
			excelhandler.GetOneData(excelhandler.PROYECT_NAME_MAIN, i)}
		//fmt.Println(newProyectMain)

		var newProyect sProyect = sProyect{
			excelhandler.GetOneData(excelhandler.PROYECT_COD, i),
			excelhandler.GetOneData(excelhandler.PROYECT_NAME, i),
			excelhandler.GetOneData(excelhandler.PROYECT_STATUS, i),
			excelhandler.GetOneData(excelhandler.PROYECT_SDATE, i)}
		//fmt.Println(newProyect)

		var newEmployee sEmployee = sEmployee{
			excelhandler.GetOneData(excelhandler.EMPLOYEE_AGRESSO, i),
			excelhandler.GetOneData(excelhandler.EMPLOYEE_NAME, i),
			excelhandler.GetOneData(excelhandler.EMPLOYEE_DOMAIN, i),
			excelhandler.GetOneData(excelhandler.EMPLOYEE_CMPYSTDATE, i),
			excelhandler.GetOneData(excelhandler.EMPLOYEE_EMAIL, i),
			excelhandler.GetOneData(excelhandler.EMPLOYEE_PC_ALTEN, i),
			excelhandler.GetOneData(excelhandler.EMPLOYEE_OFFICE, i),
			excelhandler.GetOneData(excelhandler.EMPLOYEE_LOCATION, i),
			excelhandler.GetOneData(excelhandler.EMPLOYEE_LOC_ASSOC, i),
			excelhandler.GetOneData(excelhandler.EMPLOYEE_SALARY, i),
			excelhandler.GetOneData(excelhandler.EMPLOYEE_BOND, i),
			excelhandler.GetOneData(excelhandler.EMPLOYEE_RESTAURANT, i),
			excelhandler.GetOneData(excelhandler.EMPLOYEE_INSURANCE, i),
			excelhandler.GetOneData(excelhandler.EMPLOYEE_HOLIDAYS, i),
			excelhandler.GetOneData(excelhandler.EMPLOYEE_COMMENTS, i),
			excelhandler.GetOneData(excelhandler.EMPLOYEE_SALARY_DAY, i),
			excelhandler.GetOneData(excelhandler.EMPLOYEE_BM_ADC, i),
			excelhandler.GetOneData(excelhandler.EMPLOYEE_CATEGORY, i),
			excelhandler.GetOneData(excelhandler.EMPLOYEE_N1, i),
			excelhandler.GetOneData(excelhandler.EMPLOYEE_N2, i),
			excelhandler.GetOneData(excelhandler.EMPLOYEE_INT_BM, i)}
		//fmt.Println(newEmployee)

		var completeData sCompleteData = sCompleteData{
			newClientInt,
			newClient,
			newProyectMain,
			newProyect,
			newEmployee}

		sliceExcelData = append(sliceExcelData, completeData)

		i++
	}

	for j := range sliceExcelData {
		if sliceExcelData[j].clientInt.strName == "" {
			continue
		}
		fmt.Println("------------------------------------------------")
		fmt.Println(sliceExcelData[j])
	}
}
