package main

import (
	"fmt"
	"log"

	"github.com/xuri/excelize/v2"
)

func main() {
	f, err := excelize.OpenFile("simple.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}
	if cell, err := f.GetCellValue("Sheet1", "A1"); err != nil {
		fmt.Println(err)
		return
	} else {
		fmt.Print(cell)
	}

	if err := f.SetCellValue("Sheet1", "A2", "you on99"); err != nil {
		fmt.Println(err)
		return
	}
	if saveErr := f.SaveAs("getTemplateTest.xlsx"); saveErr != nil {
		log.Fatal(saveErr)
	}
}

// func main() {
// 	http.HandleFunc("/process", process)
// 	http.ListenAndServe(":8090", nil)
// }
