package sheet

import (
	"encoding/json"
	"fmt"
	"log"

	"github.com/xuri/excelize/v2"
)

type Content struct {
	Value    string `json:"value"`
	Position string `json:"position"`
}

type StyleCell struct {
	Style      excelize.Style `json:"style"`
	TargetCell []string       `json:"targetCell"`
}

type StyleRange struct {
	Style       excelize.Style `json:"style"`
	TargetRange []Cell         `json:"targetRange"`
}

type Cell struct {
	Start string `json:"start"`
	End   string `json:"end"`
}

type MergeCell struct {
	Start string `json:"start"`
	End   string `json:"end"`
}

type RowHeight struct {
	Row    int     `json:"row"`
	Height float64 `json:"height"`
}

type ColumnWidth struct {
	StartCol string  `json:"startCol"`
	EndCol   string  `json:"endCol"`
	Width    float64 `json:"width"`
}

func ThirdSheet(f *excelize.File) {
	sheetNameThree := "Observatinos"
	f.NewSheet(sheetNameThree)

	//Cell style json stuct
	var styleArr []StyleCell
	// styleJson := `[{"Style":{"Border":[
	// 	{"Type": "left", "Color": "#FF0000", "Style": 5},
	// 	{"Type": "top", "Color": "#FF0000", "Style": 5},
	// 	{"Type": "bottom", "Color": "#000000", "Style": 5},
	// 	{"Type": "right", "Color": "#000000", "Style": 5}],"Alignment":{"Horizontal": "center"}},"TargetCell":["B1","A2","D4"]}]`

	styleJson := `[{"Style":{"Alignment":{"Horizontal": "center"},"Font":{"Underline":"single", "Bold" :true}},"TargetCell":["A5","A7"]}]`
	json.Unmarshal([]byte(styleJson), &styleArr)
	// fmt.Printf("styleArr : %+v", styleArr)

	//set up style
	for _, ranges := range styleArr {
		style, err := f.NewStyle(&ranges.Style)
		if err != nil {
			log.Print(err)
		}
		// log.Printf(" %+v", ranges)
		for _, cell := range ranges.TargetCell {
			if err := f.SetCellStyle(sheetNameThree, cell, cell, style); err != nil {
				fmt.Print(err)
				return
			}
		}
	}

	//Range style json stuct
	var styleRange []StyleRange
	styleRangeJson := `[{"Style":{"Border":[
		{"Type": "left", "Color": "#000000", "Style": 1},
		{"Type": "top", "Color": "#000000", "Style": 1},
		{"Type": "bottom", "Color": "#000000", "Style": 1},
		{"Type": "right", "Color": "#000000", "Style": 1}],"Alignment":{"Horizontal": "center","WrapText":true}},"TargetRange":[{"Start":"A9","End":"E10"}]},
		{"Style":{"Border":[
		{"Type": "left", "Color": "#000000", "Style": 1},
		{"Type": "top", "Color": "#000000", "Style": 1},
		{"Type": "bottom", "Color": "#000000", "Style": 1},
		{"Type": "right", "Color": "#000000", "Style": 1}],"Font":{"Bold" :true},"Alignment":{ "Vertical": "topLeft","WrapText":true}
		},"TargetRange":[{"Start":"A7","End":"E8"}]}
		]`

	json.Unmarshal([]byte(styleRangeJson), &styleRange)
	// fmt.Printf("styleRange : %+v", styleRange)

	//set up style
	for _, ranges := range styleRange {
		style, err := f.NewStyle(&ranges.Style)
		if err != nil {
			log.Print(err)
		}
		// log.Printf(" %+v", ranges)
		for _, cell := range ranges.TargetRange {
			if err := f.SetCellStyle(sheetNameThree, cell.Start, cell.End, style); err != nil {
				fmt.Print(err)
				return
			}
		}
	}

	// content json struct
	var contentArr []Content
	contentJson := `[
		{"value":"Monitoring of Follow-up Actoins in Safety and Environmental Walk","position":"A5"},
		{"value":"Weekly Safety Inspection dated {date} ({Weekday})","position":"A7"},
		{"value":"Item No*","position":"A8"},{"value":"Observations/Deficiencies","position":"B8"},
		{"value":"Situation Requiring  Follow-up Action","position":"C8"},
		{"value":"Recification with Record Photos","position":"D8"},
		{"value":"Follow-up Action","position":"E8"}]`
	json.Unmarshal([]byte(contentJson), &contentArr)
	for _, row := range contentArr {
		if err := f.SetCellValue(sheetNameThree, row.Position, row.Value); err != nil {
			fmt.Println(err)
			return
		}
	}

	mergeJson := `[{"start":"A5", "end":"B5"},{"start":"A7", "end":"E7"}]`
	var mergeCellRange []MergeCell
	json.Unmarshal([]byte(mergeJson), &mergeCellRange)
	// fmt.Printf(`%+v`, mergeCellRange)
	for _, ranges := range mergeCellRange {
		if err := f.MergeCell(sheetNameThree, ranges.Start, ranges.End); err != nil {
			fmt.Println(err)
			return
		}
	}

	//set Width
	widthJson := `[{"startCol":"B","endCol":"B","width":50},{"startCol":"C","endCol":"C","width":30},{"startCol":"D","endCol":"D","width":50},{"startCol":"E","endCol":"E","width":30}]`
	var widthRange []ColumnWidth
	json.Unmarshal([]byte(widthJson), &widthRange)
	// fmt.Printf(`widthRange : %+v`, widthRange)
	for _, ranges := range widthRange {
		if err := f.SetColWidth(sheetNameThree, ranges.StartCol, ranges.EndCol, ranges.Width); err != nil {
			fmt.Print(err)
			return
		}
	}

	//set Height
	heightJson := `[{"row":8,"height":30},{"row":9,"height":80},{"row":10,"height":80}]`
	var heightRange []RowHeight
	json.Unmarshal([]byte(heightJson), &heightRange)
	// fmt.Printf(`heightRange : %+v`, heightRange)
	for _, ranges := range heightRange {
		if err := f.SetRowHeight(sheetNameThree, ranges.Row, ranges.Height); err != nil {
			fmt.Print(err)
			return
		}
	}
}
