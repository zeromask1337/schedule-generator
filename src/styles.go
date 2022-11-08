package main

import "github.com/xuri/excelize/v2"

func applyStyles(f *excelize.File, sheetName string) {
	style, err := f.NewStyle(&excelize.Style{Alignment: &excelize.Alignment{Horizontal: "center", Vertical: "center", WrapText: true}})
	if err != nil {
		ErrorLogger.Println("Failed to create new style.\n", err)
	}

	if err := f.SetColStyle(sheetName, "A:AZ", style); err != nil {
		WarningLogger.Println("Failed to apply styles.\n", err)
	}

	if err := f.SetColWidth(sheetName, "A", "B", 30); err != nil {
		WarningLogger.Println("Failed to apply width to columnds\n", err)
	}
}
