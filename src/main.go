package main

import (
	"fmt"
	"github.com/xuri/excelize/v2"
	"io"
	"log"
	"net/http"
	"os"
	"path/filepath"
	"reflect"
	"sort"
	"strconv"
	"strings"
	"time"
)

var (
	InfoLogger    *log.Logger
	WarningLogger *log.Logger
	ErrorLogger   *log.Logger
)

type employee struct {
	Name        string
	Birthday    time.Time
	StartTime   time.Time
	EndTime     time.Time
	Job         string
	PhoneNumber string
	Weekend     string
}

func init() {
	// Check if args are empty
	if len(os.Args) > 1 {
		fmt.Printf("FILEPATH: %v\n", os.Args[1])
	} else {
		log.Fatal("Error: no file path argument.\n")
	}

	now := time.Now()
	hour, min, sec := now.Clock()
	year, month, day := now.Date()
	dir := filepath.Dir(os.Args[1])
	filename := fmt.Sprintf("log-%v-%v-%v_%v:%v:%v.txt", year, month, day, hour, min, sec)

	// Create logs directory
	if err := os.Mkdir(filepath.Join(dir, "logs"), os.ModePerm); err != nil {
		log.Println(err)
	}

	// Write logs file
	file, err := os.OpenFile(
		filepath.Join(dir, "logs", filename),
		os.O_APPEND|os.O_CREATE|os.O_WRONLY, 0666)
	if err != nil {
		log.Panic("Error writing logs file.\n", err)
	}
	InfoLogger = log.New(file, "[INFO] ", log.LstdFlags)
	WarningLogger = log.New(file, "[WARNING] ", log.LstdFlags)
	ErrorLogger = log.New(file, "[ERROR] ", log.LstdFlags)
}

func main() {
	desiredTime := time.Now().AddDate(0, 1, 0)
	nextMonth := buildDate(desiredTime.Year(), desiredTime.Month(), 1)
	year, month, _ := nextMonth.Date()
	daysInMonth := daysIn(month, year)
	daysOfWeek := [7]string{"вс", "пн", "вт", "ср", "чт", "пт", "сб"}

	/* SCAN EMPLOYEE DATABASE SHEET */

	// Import file
	filePath := os.Args[1]
	f, err := excelize.OpenFile(filePath)
	if err != nil {
		ErrorLogger.Fatal("Can't open excel file. ", err)
	}

	defer func() {
		if err := f.Save(); err != nil {
			ErrorLogger.Fatal("Can't save file. ", err)
		}
	}()

	// Get all the rows in the database.
	rows, err := f.GetRows("Сотрудники")
	if err != nil {
		ErrorLogger.Fatal("Can't get sheet contents. ", err)
	}

	/* CREATE EMPLOYEES COLLECTION */

	var (
		employees []employee
		headers   = rows[0]
	)

	for i, row := range rows {
		if i == 0 {
			continue
		}
		e := new(employee)
		metaValue := reflect.ValueOf(e).Elem()

		// Build employee
		for j, name := range headers {
			field := metaValue.FieldByName(name)
			if field.IsValid() && field.CanSet() {
				switch name {
				case "Birthday":
					layout := "01-02-06"
					res, err := time.Parse(layout, row[j])
					if err != nil {
						ErrorLogger.Fatalf("Error parsing Birthday of %v: %v", e.Name, err)
					}
					field.Set(reflect.ValueOf(res))

				case "StartTime", "EndTime":
					floatTime, err := strconv.ParseFloat(row[j], 64)
					if err != nil {
						ErrorLogger.Fatalf("Error parsing float string of %v: %v", e.Name, err)
					}
					timeClock, err := excelize.ExcelDateToTime(floatTime, false)
					if err != nil {
						ErrorLogger.Fatalf("Error converting float to clock of %v: %v", e.Name, err)
					}
					field.Set(reflect.ValueOf(timeClock))
				default:
					field.SetString(row[j])
				}
			} else {
				WarningLogger.Println("Field %s does not exist in struct", name)
			}
		}
		if e.StartTime.Hour() > e.EndTime.Hour() {
			newDate := e.EndTime.AddDate(0, 0, 1)
			metaValue.FieldByName("EndTime").Set(reflect.ValueOf(newDate))
		}
		employees = append(employees, *e)
		InfoLogger.Printf("Added %v to collection %+v", e.Name, e)
	}

	/* FILL SPREADSHEET WITH CALENDAR DAYS, EMPLOYEES AND SCHEDULE */

	var (
		sheetName    = fmt.Sprintf("%v %v", month, year)
		sheetIndex   = f.NewSheet(sheetName)
		weekDaySlice []string
		weekDaysMap  = map[int]int{}
		monthDays    []int
		monthRow     = []any{"ФИО/Дата"}
	)
	InfoLogger.Printf("Created sheet %v: %v\n", sheetIndex, sheetName)

	// Set row with calendar days
	for i := 1; i <= daysInMonth; i++ {
		wdi := int(buildDate(desiredTime.Year(), desiredTime.Month(), i).Weekday())
		weekDaysMap[i] = wdi
		weekDaySlice = append(weekDaySlice, daysOfWeek[weekDaysMap[i]])
	}
	monthDays = getKeys(weekDaysMap)
	sort.Ints(monthDays)

	// Append month days checkout column headers
	for _, v := range monthDays {
		monthRow = append(monthRow, v)
	}
	monthRow = append(
		monthRow,
		"Норма часов, согласно производственному календарю",
		"Отработано в месяц (часов)",
		"Подпись работника",
	)

	// Insert calendar row
	cellMR, err := excelize.CoordinatesToCellName(2, 5) // TODO: Maybe refactor into method ?
	if err != nil {
		ErrorLogger.Fatal("Creating cell from coordinates failed")
	}
	if err := f.SetSheetRow(sheetName, cellMR, &monthRow); err != nil {
		ErrorLogger.Fatalf("Setting sheet row %v failed", cellMR, err)
	}

	// Insert week day row
	cellWDS, err := excelize.CoordinatesToCellName(3, 6)
	if err != nil {
		ErrorLogger.Fatal("Creating cell from coordinates failed")
	}
	if err := f.SetSheetRow(sheetName, cellWDS, &weekDaySlice); err != nil {
		ErrorLogger.Fatal("Setting sheet row %v failed", cellWDS, err)
	}

	/* SET ROWS WITH EMPLOYEES */

	applyGeneralStyles(f, sheetName, daysInMonth, employees)

	//Fetch weekends and holidays from isdayoff API
	res, err := http.Get(fmt.Sprintf("https://isdayoff.ru/api/getdata?year=%v&month=%v&cc=ru&pre=1&covid=0&sd=0", year, int(month)))
	if err != nil {
		ErrorLogger.Fatalln("Fetching data from 'isdayoff' API failed.\n", err)
	} else {
		InfoLogger.Println("Success fetch from 'isdayoff' API")
	}

	data, _ := io.ReadAll(res.Body)
	dataString := string(data)

	cursor := 7
	for i, e := range employees {
		worktimeRow := []string{e.Job, e.Name}
		totalHoursRow := []any{}

		for j, ch := range dataString {
			date := buildDate(year, month, j+1)
			wd := int(date.Weekday())
			weekend := toInt(strings.Split(e.Weekend, ""))

			switch string(ch) {
			case "0", "4":
				start := e.StartTime.Format("15:04")
				end := e.EndTime.Format("15:04")
				cellValue := fmt.Sprintf("%v-%v", start, end)

				if e.Birthday.Month() == date.Month() && e.Birthday.Day() == date.Day() {
					cellValue += ", ДР"
					x, y := buildCoordinates(j+3, 7+i*2, j+3, 7+i*2)
					paintBirthday(f, sheetName, x, y)
				}

				isWeekend := false
				for _, v := range weekend {
					if weekDaysMap[j+1] == v {
						isWeekend = true
					}
				}
				if isWeekend == true {
					worktimeRow = append(worktimeRow, "B")
					totalHoursRow = append(totalHoursRow, "в")
				} else {
					worktimeRow = append(worktimeRow, cellValue)
					workDuration := e.EndTime.Sub(e.StartTime) - time.Hour*1 // lunch
					totalHoursRow = append(totalHoursRow, workDuration.Hours())
				}

			case "1":
				var cellValue string
				if wd == 6 || wd == 0 {
					start := e.StartTime.Format("15:04")
					end := e.EndTime.Format("15:04")
					cellValue = fmt.Sprintf("%v-%v", start, end)
					st, ec := buildCoordinates(j+3, 5, j+3, 6+len(employees)*2)
					paintWeekend(f, sheetName, st, ec)
				} else {
					cellValue = "ПРАЗДНИК"
					x, y := buildCoordinates(j+3, 5, j+3, 6+len(employees)*2)
					paintHoliday(f, sheetName, x, y)
				}
				if e.Birthday.Month() == date.Month() && e.Birthday.Day() == date.Day() {
					cellValue += ", ДР"
					x, y := buildCoordinates(j+3, 7+i*2, j+3, 7+i*2)
					paintBirthday(f, sheetName, x, y)
				}

				isWeekend := false
				for _, v := range weekend {
					if weekDaysMap[j+1] == v {
						isWeekend = true
					}
				}
				if isWeekend == true {
					worktimeRow = append(worktimeRow, "B")
					totalHoursRow = append(totalHoursRow, "в")
				} else if cellValue == "ПРАЗДНИК" || cellValue == "ПРАЗДНИК, ДР" {
					worktimeRow = append(worktimeRow, cellValue)
					totalHoursRow = append(totalHoursRow, "в")
				} else {
					worktimeRow = append(worktimeRow, cellValue)
					workDuration := e.EndTime.Sub(e.StartTime) - time.Hour*1 // lunch
					totalHoursRow = append(totalHoursRow, workDuration.Hours())
				}

			case "2":
				start := e.StartTime.Format("15:04")
				end := e.EndTime.Format("15:04")
				cellValue := fmt.Sprintf("%v-%v%v", start, end, ", СОКР")

				x, y := buildCoordinates(j+3, 5, j+3, 6+len(employees)*2)
				paintHalfDay(f, sheetName, x, y)

				if e.Birthday.Month() == date.Month() && e.Birthday.Day() == date.Day() {
					cellValue += ", ДР"
					x, y := buildCoordinates(j+3, 7+i*2, j+3, 7+i*2)
					paintBirthday(f, sheetName, x, y)
				}

				isWeekend := false
				for _, v := range weekend {
					if weekDaysMap[j+1] == v {
						isWeekend = true
					}
				}
				if isWeekend == true {
					worktimeRow = append(worktimeRow, "B")
					totalHoursRow = append(totalHoursRow, "в")
				} else {
					worktimeRow = append(worktimeRow, cellValue)
					workDuration := e.EndTime.Sub(e.StartTime) - time.Hour*2 // lunch
					totalHoursRow = append(totalHoursRow, workDuration.Hours())
				}
			}
		}

		// Inserting and merging rows
		totalHoursRow = append(totalHoursRow, "")
		if err := f.SetSheetRow(sheetName, fmt.Sprintf("A%v", cursor), &worktimeRow); err != nil {
			ErrorLogger.Fatal("Inserting worktimeRow on A failed. ", err)
		}
		if err := f.SetSheetRow(sheetName, fmt.Sprintf("C%v", cursor+1), &totalHoursRow); err != nil {
			ErrorLogger.Fatal("Inserting totalhoursRow on C failed. ", err)
		}

		formulaXY, err := excelize.CoordinatesToCellName(2+len(weekDaySlice)+2, cursor+1)
		if err != nil {
			WarningLogger.Println("Failed to convert coordinates to cellname.\n", err)
		}
		frX, _ := excelize.CoordinatesToCellName(3, cursor+1)
		frY, _ := excelize.CoordinatesToCellName(2+len(weekDaySlice), cursor+1)
		formulaRange := fmt.Sprintf("=SUM(%v:%v)", frX, frY)
		if err := f.SetCellFormula(sheetName, formulaXY, formulaRange); err != nil {
			WarningLogger.Println("Failed to insert cell formula.\n", err)
		}
		if err := f.MergeCell(sheetName, fmt.Sprintf("A%v", cursor), fmt.Sprintf("A%v", cursor+1)); err != nil {
			ErrorLogger.Fatal("Merging cell A failed. ", err)
		}
		if err := f.MergeCell(sheetName, fmt.Sprintf("B%v", cursor), fmt.Sprintf("B%v", cursor+1)); err != nil {
			ErrorLogger.Fatal("Merging cell B failed. ", err)
		}
		cursor += 2
	}
}

// getKeys returns map keys
func getKeys(m map[int]int) (keys []int) {
	for k := range m {
		keys = append(keys, k)
	}
	return keys
}

// toInt returns int slice
func toInt(s []string) (slice []int) {
	for _, elem := range s {
		if convertedStr, err := strconv.Atoi(elem); err != nil {
			log.Printf("toInt Error %v: %v", elem, err)
		} else {
			slice = append(slice, convertedStr)
		}
	}
	return slice
}

// buildDate returns Time type using only year, month and day
func buildDate(year int, month time.Month, day int) time.Time {
	return time.Date(year, month, day, 0, 0, 0, 0, time.UTC)
}

// daysIn returns number of days in specified month
func daysIn(m time.Month, year int) int {
	return buildDate(year, m+1, 0).Day()
}

// buildCoordinates returns x and y of input coordinates
func buildCoordinates(collX int, rowX int, colY int, rowY int) (startCell string, endCell string) {
	st, err := excelize.CoordinatesToCellName(collX, rowX)
	if err != nil {
		ErrorLogger.Println("Failed to make cellname from coordinates.\n", err)
	}
	ec, err := excelize.CoordinatesToCellName(colY, rowY)
	if err != nil {
		ErrorLogger.Println("Failed to make cellname from coordinates.\n", err)
	}
	return st, ec
}
