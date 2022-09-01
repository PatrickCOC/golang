package sheet

import (
	"fmt"
	"strconv"

	"github.com/xuri/excelize/v2"
)

type Checklist struct {
	qid      string
	question string
}

func FirstSheet(f *excelize.File) {
	sheetNameOne := "Checklist"
	f.SetSheetName("Sheet1", sheetNameOne)

	box, _ := f.NewStyle(&excelize.Style{
		Border: []excelize.Border{
			{Type: "left", Color: "#000000", Style: 5},
			{Type: "top", Color: "#000000", Style: 5},
			{Type: "bottom", Color: "#000000", Style: 5},
			{Type: "right", Color: "#000000", Style: 5},
		}, Alignment: &excelize.Alignment{
			Horizontal: "center",
		},
	})

	bottomLine, _ := f.NewStyle(&excelize.Style{
		Border: []excelize.Border{
			{Type: "bottom", Color: "#000000", Style: 1},
		},
	})

	underline, _ := f.NewStyle(&excelize.Style{
		Font: &excelize.Font{Underline: "single"},
	})

	style, _ := f.NewStyle(&excelize.Style{
		Alignment: &excelize.Alignment{Horizontal: "center"},
	})

	autoChangeLine, _ := f.NewStyle(
		&excelize.Style{
			Alignment: &excelize.Alignment{
				Vertical: "topLeft",
				WrapText: true,
			},
		},
	)
	// insert cell value
	data := [][]interface{}{
		{"WEEKLY CONSTRUCTION SAFETY INSPECTION CHECK-LIST"},
		{"Contract No."},
		{"Contract Title"},
		{"Date", "", "Time"},
		{"Contractor Representative"},
		{"Consultant Representative"},
	}
	rowNum := 1
	for i, row := range data {
		rowNum++
		startCell, err := excelize.JoinCellName("A", rowNum)
		if err != nil {
			fmt.Println(err)
			return
		}
		if i == 0 {
			startCell, err = excelize.JoinCellName("A", i+1)
			f.SetCellStyle(sheetNameOne, "A1", "P1", style)
			if err != nil {
				fmt.Println(err)
				return
			}
		} else if i == 1 {
			startCell, err = excelize.JoinCellName("A", i+2)
			f.SetCellStyle(sheetNameOne, "A"+strconv.Itoa(rowNum+1), "B"+strconv.Itoa(rowNum+1), bottomLine)
			if err != nil {
				fmt.Println(err)
				return
			}
		} else if i < 7 && i > 1 {
			rowNum = rowNum + 1
			startCell, err = excelize.JoinCellName("A", rowNum)
			if i < 6 {
				f.SetCellStyle(sheetNameOne, "A"+strconv.Itoa(rowNum+1), "B"+strconv.Itoa(rowNum+1), bottomLine)
			}
			if err != nil {
				fmt.Println(err)
				return
			}
		}
		if err := f.SetSheetRow(sheetNameOne, startCell, &row); err != nil {
			fmt.Println(err)
			return
		}
	}

	checklistTitle := [][]interface{}{
		{"General Site Safety Hazard"},
		{"Housekeeping", "", "A", "", "B", "", "C", "", "D", "", "N/A"},
	}
	rowNum += 1
	for _, row := range checklistTitle {
		rowNum++
		startCell, err := excelize.JoinCellName("A", rowNum+1)
		if err != nil {
			fmt.Println(err)
			return
		}
		if err := f.SetSheetRow(sheetNameOne, startCell, &row); err != nil {
			fmt.Println(err)
			return
		}
		f.SetCellStyle(sheetNameOne, "A"+strconv.Itoa(rowNum+1), "B"+strconv.Itoa(rowNum+1), underline)
	}

	mergeCellRange := [][]string{{"A1", "P1"}, {"A3", "B3"}, {"A5", "B5"}, {"A9", "B9"}, {"A13", "B13"}, {"A14", "B14"}}

	for _, ranges := range mergeCellRange {
		if err := f.MergeCell(sheetNameOne, ranges[0], ranges[1]); err != nil {
			fmt.Println(err)
			return
		}
	}

	checklistabc := []Checklist{
		{qid: "A.1.1", question: "Works area kept clean, tidy and in good hygiene condition"},
		{qid: "A.1.2", question: "Materials stocked/stored in orderly manner and at designated area, fenced off and separated from temporary access, not stacked in unstable condition, cylindrical materials wedged"},
		{qid: "A.1.3", question: "Waste materials and debris contained appropriately, and cleared off regularly"},
		{qid: "A.1.4", question: "Chemicals / oils properly stored to avoid direct sunlight, prevent spillage,storage not exceeding statutory limit"},
		{qid: "A.1.5", question: "Sufficient fire extinguishers provided"},
		{qid: "A.1.6", question: "Fire Assembly Point provided"},
		{qid: "A.1.7", question: "First aid facility (i.e. first aid box) adequately provided"},
		{qid: "A.1.8", question: "Welfare facility (i.e. rest shelter) adequately provided and maintained in good"},
	}

	for _, row := range checklistabc {
		rowNum += 2

		startCell, err := excelize.JoinCellName("A", rowNum)
		if row.qid != "" {
			if err != nil {
				fmt.Println(err)
				return
			}
			if err := f.SetCellValue(sheetNameOne, startCell, row.qid); err != nil {
				fmt.Print("row.qid err:", err)
				return
			}
		}

		secondStartCell, err := excelize.JoinCellName("B", rowNum)
		if row.question != "" {
			if err != nil {
				fmt.Println(err)
				return
			}
			if err := f.SetCellValue(sheetNameOne, secondStartCell, row.question); err != nil {
				fmt.Println("row.qusetion err:", err)
				return
			}
		}

		tooLong := len(row.question) / 20
		if tooLong > 1 {
			if err := f.SetRowHeight(sheetNameOne, rowNum, float64(tooLong*4+10)); err != nil {
				fmt.Print(err)
				return
			}
		}
		f.SetCellStyle(sheetNameOne, "A"+strconv.Itoa(rowNum), "B"+strconv.Itoa(rowNum), autoChangeLine)
		f.SetCellStyle(sheetNameOne, "C"+strconv.Itoa(rowNum), "C"+strconv.Itoa(rowNum), box)
		f.SetCellStyle(sheetNameOne, "E"+strconv.Itoa(rowNum), "E"+strconv.Itoa(rowNum), box)
		f.SetCellStyle(sheetNameOne, "G"+strconv.Itoa(rowNum), "G"+strconv.Itoa(rowNum), box)
		f.SetCellStyle(sheetNameOne, "I"+strconv.Itoa(rowNum), "I"+strconv.Itoa(rowNum), box)
		f.SetCellStyle(sheetNameOne, "K"+strconv.Itoa(rowNum), "K"+strconv.Itoa(rowNum), box)
	}
	if err := f.SetColWidth(sheetNameOne, "A", "M", 10); err != nil {
		fmt.Print(err)
		return
	}

	if err := f.SetColWidth(sheetNameOne, "B", "B", 80); err != nil {
		fmt.Print(err)
		return
	}

}
