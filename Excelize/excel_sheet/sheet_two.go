package sheet

import (
	"fmt"
	h "go/excel/helper"
	"reflect"
	"strconv"

	"github.com/xuri/excelize/v2"
)

type BlockLetter struct {
	itemNumber int
	name       string
	title      string
	signatures string
}

type FreeGuy struct {
	column_1 string
	column_2 string
	column_3 string
	column_4 string
	column_5 string
	column_6 string
}

func SecondSheet(f *excelize.File) {
	sheetNameTwo := "Summary"
	f.NewSheet(sheetNameTwo)
	underline, _ := f.NewStyle(&excelize.Style{
		Font: &excelize.Font{Underline: "single"},
	})
	header := [][]interface{}{
		{"Weekly Safety Walk Inspection Report"},
		{"Summary of Follow-up Actions"},
		{""},
		{"Part I:"},
		{"Contract No.:", "", "", "Contract Title:", ""},
		{"Date of Inspection:", "", "", "Time", ""},
		{"Consultant Representative"},
	}
	rowNum := 1
	for _, row := range header {
		rowNum++
		startCell, err := excelize.JoinCellName("A", rowNum)
		if err != nil {
			fmt.Println(err)
			return
		}
		if err := f.SetSheetRow(sheetNameTwo, startCell, &row); err != nil {
			fmt.Println(err)
			return
		}
	}
	rowNum += 1
	blockLetter := [][]interface{}{
		{"", "Name in Block Letters", "", "Designated", "", "Organization", "", "Name in Block Letters", "", "Designated", "", "Organization"},
	}

	border, _ := f.NewStyle(&excelize.Style{
		Border: []excelize.Border{
			{Type: "left", Color: "#000000", Style: 1},
			{Type: "top", Color: "#000000", Style: 1},
			{Type: "bottom", Color: "#000000", Style: 1},
			{Type: "right", Color: "#000000", Style: 1},
		}, Alignment: &excelize.Alignment{
			Horizontal: "center",
		},
	})

	for _, row := range blockLetter {
		rowNum++
		startCell, err := excelize.JoinCellName("A", rowNum)
		if err != nil {
			fmt.Println(err)
			return
		}
		if err := f.SetSheetRow(sheetNameTwo, startCell, &row); err != nil {
			fmt.Println(err)
			return
		}
	}

	letterData := []BlockLetter{
		{name: "Joe", title: "Admin", signatures: "signatures"},
		{name: "Tom", title: "PM", signatures: "signatures"},
		{name: "Ken", title: "PM", signatures: "signatures"},
		{name: "Peter", title: "PM", signatures: "signatures"},
		{name: "Sam", title: "PM", signatures: "signatures"},
		{name: "Tony", title: "Inspector", signatures: "signatures"},
		{name: "Ben", title: "Engineer", signatures: "signatures"},
	}

	for idx := range letterData {
		letterData[idx].itemNumber = idx + 1
		// num := []interface{}{idx + 1}
		// letterData[idx] = append(num, letterData[idx]...)
	}

	// blockLetter = append(blockLetter, letterData...)
	secondRow := rowNum
	for i, row := range letterData {
		if i > 4 {
			secondRow++
			numCell, _ := excelize.JoinCellName("G", secondRow)
			nameCell, _ := excelize.JoinCellName("H", secondRow)
			titleCell, _ := excelize.JoinCellName("J", secondRow)
			signatureCell, _ := excelize.JoinCellName("L", secondRow)
			if err := f.SetCellValue(sheetNameTwo, numCell, row.itemNumber); err != nil {
				fmt.Println(err)
				return
			}
			if err := f.SetCellValue(sheetNameTwo, nameCell, row.name); err != nil {
				fmt.Println(err)
				return
			}
			f.SetCellStyle(sheetNameTwo, nameCell, nameCell, underline)
			if err := f.SetCellValue(sheetNameTwo, titleCell, row.title); err != nil {
				fmt.Println(err)
				return
			}
			f.SetCellStyle(sheetNameTwo, titleCell, titleCell, underline)
			if err := f.SetCellValue(sheetNameTwo, signatureCell, row.signatures); err != nil {
				fmt.Println(err)
				return
			}
			f.SetCellStyle(sheetNameTwo, signatureCell, signatureCell, underline)

		} else {
			rowNum++

			numCell, _ := excelize.JoinCellName("A", rowNum)
			nameCell, _ := excelize.JoinCellName("B", rowNum)
			titleCell, _ := excelize.JoinCellName("D", rowNum)
			signatureCell, _ := excelize.JoinCellName("F", rowNum)

			if err := f.SetCellValue(sheetNameTwo, numCell, row.itemNumber); err != nil {
				fmt.Println(err)
				return
			}
			if err := f.SetCellValue(sheetNameTwo, nameCell, row.name); err != nil {
				fmt.Println(err)
				return
			}
			f.SetCellStyle(sheetNameTwo, nameCell, nameCell, underline)
			if err := f.SetCellValue(sheetNameTwo, titleCell, row.title); err != nil {
				fmt.Println(err)
				return
			}
			f.SetCellStyle(sheetNameTwo, titleCell, titleCell, underline)
			if err := f.SetCellValue(sheetNameTwo, signatureCell, row.signatures); err != nil {
				fmt.Println(err)
				return
			}
			f.SetCellStyle(sheetNameTwo, signatureCell, signatureCell, underline)
		}

		f.SetCellStyle(sheetNameTwo, "G"+strconv.Itoa(rowNum+1), "G"+strconv.Itoa(rowNum+1), underline)

	}
	if err := f.SetColWidth(sheetNameTwo, "A", "M", 14); err != nil {
		fmt.Print(err)
		return
	}
	rowNum += 1
	tableHeader := []FreeGuy{
		{column_1: "Item No.", column_2: "Location", column_3: "Situation Requiring Follow-up action", column_4: "Agreed Due Date for Completion", column_5: "Date Completed", column_6: "Remarks"},
	}
	tableData := []FreeGuy{
		{column_1: "string1", column_2: "string2", column_3: "string3", column_4: "string4", column_5: "string5", column_6: "string6"},
		{column_1: "string1", column_2: "string2", column_3: "string3", column_4: "string4", column_5: "string5", column_6: "string6"},
		{column_1: "string1", column_2: "string2", column_3: "string3", column_4: "string4", column_5: "string5", column_6: "string6"},
		{column_1: "string1", column_2: "string2", column_3: "string3", column_4: "string4", column_5: "string5", column_6: "string6"},
		{column_1: "string1", column_2: "string2", column_3: "string3", column_4: "string4", column_5: "string5", column_6: "string6"},
		{column_1: "string1", column_2: "string2", column_3: "string3", column_4: "string4", column_5: "string5", column_6: "string6"},
		{column_1: "string1", column_2: "string2", column_3: "string3", column_4: "string4", column_5: "string5", column_6: "string6"},
		{column_1: "string1", column_2: "string2", column_3: "string3", column_4: "string4", column_5: "string5", column_6: "string6"},
	}
	tableHeader = append(tableHeader, tableData...)
	f.SetCellStyle(sheetNameTwo, "A"+strconv.Itoa(rowNum), "M"+strconv.Itoa(len(tableHeader)+rowNum), border)
	//create table
	// if err := f.AddTable(sheetNameTwo, "A"+strconv.Itoa(rowNum), "M"+strconv.Itoa(len(tableHeader)+rowNum),
	// 	`{
	// 		"table_name": "table",
	// 		"table_style": "TableStyleMedium2",
	// 		"show_first_column": true,
	// 		"show_last_column": true,
	// 		"show_row_stripes": false,
	// 		"show_column_stripes": true
	// 	}`); err != nil {
	// 	fmt.Print(err)
	// 	return
	// }
	alphabet := h.ToInt([]rune("A"))
	for _, row := range tableHeader {
		column_Cell_x, _ := excelize.JoinCellName(h.ToChar(alphabet), rowNum)
		// column_Cell_y, _ := excelize.JoinCellName(h.ToChar(alphabet+1), rowNum)
		// x := h.ToChar(alphabet) + strconv.Itoa(rowNum)
		// y := h.ToChar(alphabet+1) + strconv.Itoa(rowNum)
		// if err := f.MergeCell(sheetNameTwo, x, y); err != nil {
		// 	fmt.Println(err)
		// 	return
		// }
		v := reflect.ValueOf(row)
		values := make([]interface{}, reflect.TypeOf(FreeGuy{}).NumField())
		for i := 0; i < reflect.TypeOf(FreeGuy{}).NumField(); i++ {
			values[i] = v.Field(i)
		}
		rowNum++
		if err := f.SetSheetRow(sheetNameTwo, column_Cell_x, &values); err != nil {
			fmt.Println(err)
			return
		}
	}
}
