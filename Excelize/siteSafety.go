package main

import (
	sheet "go/excel/excel_sheet"
	_ "image/gif"
	_ "image/jpeg"
	_ "image/png"
	"log"

	"github.com/xuri/excelize/v2"
)

func main() {
	f := excelize.NewFile()
	sheet.FirstSheet(f)
	sheet.SecondSheet(f)
	sheet.ThirdSheet(f)

	if saveErr := f.SaveAs("siteSafety.xlsx"); saveErr != nil {
		log.Fatal(saveErr)
	}
}
