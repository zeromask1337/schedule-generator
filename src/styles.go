package main

import (
	"github.com/xuri/excelize/v2"
)

func applyGeneralStyles(f *excelize.File, sheetName string, dim int, employees []employee) {
	style, err := f.NewStyle(&excelize.Style{
		Alignment: &excelize.Alignment{
			Horizontal: "center",
			Vertical:   "center",
			WrapText:   true,
		},
		Border: []excelize.Border{
			{
				Type:  "top",
				Color: "#000000",
				Style: 1,
			},
			{
				Type:  "right",
				Color: "#000000",
				Style: 1,
			},
			{
				Type:  "bottom",
				Color: "#000000",
				Style: 1,
			},
			{
				Type:  "left",
				Color: "#000000",
				Style: 1,
			},
		},
	})
	if err != nil {
		ErrorLogger.Println("Failed to create new style.1\n", err)
	}

	startCell, err := excelize.CoordinatesToCellName(1, 5)
	if err != nil {
		ErrorLogger.Println("Failed to create cellname from coordinates\n", err)
	}
	endCell, err := excelize.CoordinatesToCellName(2+dim+3, 6+len(employees)*2) // 2 - left offset, 3 - right offset
	if err != nil {
		ErrorLogger.Println("Failed to create cellname from coordinates\n", err)
	}

	if err := f.SetCellStyle(sheetName, startCell, endCell, style); err != nil {
		WarningLogger.Println("Failed to apply styles.\n", err)
	}
	if err := f.SetColWidth(sheetName, "A", "B", 30); err != nil {
		WarningLogger.Println("Failed to apply width to columns\n", err)
	}
	if err := f.SetColWidth(sheetName, "C", "AJ", 11); err != nil {
		WarningLogger.Println("Failed to apply width to columns\n", err)
	}
}

func paintWeekend(file *excelize.File, sheetName string, startCell, endCell string) {
	style, err := file.NewStyle(&excelize.Style{
		Alignment: &excelize.Alignment{
			Horizontal: "center",
			Vertical:   "center",
			WrapText:   true,
		},
		Fill: excelize.Fill{
			Type:    "pattern",
			Color:   []string{"#FFFF00"},
			Pattern: 1},
		Border: []excelize.Border{
			{
				Type:  "top",
				Color: "#000000",
				Style: 1,
			},
			{
				Type:  "right",
				Color: "#000000",
				Style: 1,
			},
			{
				Type:  "bottom",
				Color: "#000000",
				Style: 1,
			},
			{
				Type:  "left",
				Color: "#000000",
				Style: 1,
			},
		},
	})
	if err != nil {
		ErrorLogger.Println("Failed to create new style.2\n", err)
	}

	if err := file.SetCellStyle(sheetName, startCell, endCell, style); err != nil {
		WarningLogger.Println("Failed to apply color weekend.\n", err)
	}
}
