package main

import (
	"fmt"
	"io"
	"log"
	"net/http"
	"os"
	"strings"

	_ "image/gif"
	_ "image/jpeg"
	_ "image/png"

	"github.com/xuri/excelize/v2"
)

type input struct {
	value string
}

func downloadImage(url string) string {

	// don't worry about errors
	response, e := http.Get(url)
	name := strings.Split(url, "signatures/")
	if e != nil {
		log.Fatal(e)
	}
	defer response.Body.Close()

	fileName := name[len(name)-1] + ".jpeg"
	//open a file for writing
	file, err := os.Create(fileName)
	if err != nil {
		log.Fatal(err)
	}

	defer file.Close()

	// Use io.Copy to just dump the response body to the file. This supports huge files
	_, err = io.Copy(file, response.Body)
	if err != nil {
		log.Fatal(err)
	}
	fmt.Println("Success!")
	return fileName
}

func removeImage(name string) {

	err := os.Remove(name)
	if err != nil {
		log.Fatal(err)
	}
	fmt.Println("Remove Success!")
}

func simple() {

	f := excelize.NewFile()
	sheetName := "risc"
	f.SetSheetName("Sheet1", sheetName)

	// insert cell value
	data := [][]interface{}{
		{"testing"},
		{"item no", "location", "protion", "activity", "labour", "plant", "remarks"},
		{"trade", "code", "no.", "working hour", "type", "id", "working", "idle", "working hour", "plant ovnership"},
		{"AA", "BB", 3, 4, 5, 6, 7, 8, 9},
		{"AA1", "BB8", 3, 4, 5, 6, 7, 8, 9},
		{"AA2", "BB7", 3, 4, 5, 6, 7, 8, 9},
		{"AA3", "BB6", 3, 4, 5, 6, 7, 8, 9},
		{"AA4", "BB5", 3, 4, 5, 6, 7, 8, 9},
	}

	for i, row := range data {
		startCell, err := excelize.JoinCellName("A", i+1)
		if err != nil {
			fmt.Println(err)
			return
		}
		// str := fmt.Sprintf("%v", row[0])
		a, _ := row[0].([]interface{})
		if err := f.SetSheetRow(sheetName, startCell, &a); err != nil {
			fmt.Println(err)
			return
		}
	}

	// excel formula
	formulaType, ref := excelize.STCellFormulaTypeShared, "B10:B15"
	if err := f.SetCellFormula(sheetName, "A14", "=SUM(B4:D4)", excelize.FormulaOpts{Ref: &ref, Type: &formulaType}); err != nil {
		fmt.Print(err)
		return
	}

	//merge cell
	mergeCellRange := [][]string{{"A1", "A2"}, {"B1", "B2"}, {"C1", "C2"}, {"D1", "D2"}}

	for _, ranges := range mergeCellRange {
		if err := f.MergeCell(sheetName, ranges[0], ranges[1]); err != nil {
			fmt.Println(err)
			return
		}
	}

	//create table
	if err := f.AddTable(sheetName, "A3", "J8",
		`{
		"table_name": "table",
		"table_style": "TableStyleLight2"
		}`); err != nil {
		fmt.Print(err)
		return
	}

	// add filter
	// f.AutoFilter(sheetName, "A1", "J8", "")

	//style setting
	style, _ := f.NewStyle(&excelize.Style{
		Border: []excelize.Border{
			{Type: "left", Color: "#000000", Style: 1},
			{Type: "top", Color: "#000000", Style: 1},
			{Type: "bottom", Color: "#000000", Style: 1},
			{Type: "right", Color: "#000000", Style: 1},
		}, Alignment: &excelize.Alignment{
			Horizontal: "center",
		},
	})
	f.SetCellStyle(sheetName, "A1", "J8", style)

	styleHeader, err := f.NewStyle(&excelize.Style{
		Alignment: &excelize.Alignment{Horizontal: "center"},
		Fill:      excelize.Fill{Type: "pattern", Color: []string{"#DFEBF6"}, Pattern: 1},
	})
	if err != nil {
		fmt.Println(err)
		return
	}
	if f.SetCellStyle(sheetName, "A1", "A1", styleHeader); err != nil {
		fmt.Println(err)
		return
	}

	//set cell width
	if err := f.SetColWidth(sheetName, "A", "J", 14); err != nil {
		fmt.Print(err)
		return
	}

	// add chart
	if err := f.AddChart(sheetName, "A20", `{
		"type": "col",
		"series": [
			{
			"name": "risc!$A$3",
			"categories": "risc!$B$3:$I$3",
			"values": "risc!$B$4:$I$4"
			}
		],
		"format": {
			"x_scale": 1.5,
			"x_offset": 10,
			"y_offset": 20
		}
	}`); err != nil {
		fmt.Print(err)
		return
	}

	// Decode base64 image
	// b64data := url[strings.IndexByte(url, ',')+1:]
	// sDec, err1 := b64.StdEncoding.DecodeString(b64data)
	// log.Print(err1)
	// if err := f.AddPictureFromBytes(sheetName, "A2", "", "Excel Logo", ".jpg", sDec); err != nil {
	// 	fmt.Println(err)
	// }

	// If Image is url
	url := "s"
	if name := downloadImage(url); name != "" {
		if err := f.AddPicture(sheetName, "A10", name, `{
			"x_offset": 15,
			"y_offset": 15,
			"x_scale": 0.2,
			"y_scale": 0.2
		}`); err != nil {
			fmt.Println(err)
		}
		removeImage(name)
		if err != nil {
			fmt.Println(err)
		}
	}

	//ShowGridLines
	if err := f.SetSheetViewOptions(sheetName, 0, excelize.ShowGridLines(false)); err != nil {
		fmt.Print(err)
		return
	}

	//create and remvoe freeze panes and split panes
	if err = f.SetPanes(sheetName, `{
		"freeze":true,
		"split": false,
		"x_split": 0,
		"y_split": 3,
		"top_left_cell": "A4",
		"active_pane": "bottomLeft"
	}`); err != nil {
		fmt.Print(err)
		return
	}

	if saveErr := f.SaveAs("excel_test.xlsx"); saveErr != nil {
		log.Fatal(saveErr)
	}
}
