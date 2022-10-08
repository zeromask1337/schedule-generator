package main

import (
	"fmt"
	"github.com/xuri/excelize/v2"
	"log"
	"os"
	"reflect"
	"sort"
	"strconv"
	"strings"
	"time"
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

// daysIn returns number of days in specified month
func daysIn(m time.Month, year int) int {
	return time.Date(year, m+1, 0, 0, 0, 0, 0, time.UTC).Day()
}

func main() {
	desiredTime := time.Now().AddDate(0, 1, 0)
	nextMonth := time.Date(desiredTime.Year(), desiredTime.Month(), 1, 0, 0, 0, 0, time.UTC)
	year, month, _ := nextMonth.Date()
	daysInMonth := daysIn(month, year)
	daysOfWeek := [7]string{"вс", "пн", "вт", "ср", "чт", "пт", "сб"}

	// Check if args are empty
	if len(os.Args) > 1 {
		log.Printf("[INFO] File path: %v", os.Args[1])
	} else {
		log.Panicln("[ERROR] no file path argument")
	}
	filePath := os.Args[1]

	f, err := excelize.OpenFile(filePath)
	if err != nil {
		fmt.Println(err)
		return
	}
	defer func() {
		if err := f.Save(); err != nil {
			fmt.Println(err)
		}
	}()

	// Get all the rows in the Sheet1.
	rows, err := f.GetRows("Сотрудники")
	if err != nil {
		log.Panicln(err)
	}

	// Create employees collection
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
						log.Fatalf("Error parsing Birthday of %v: %v", e.Name, err)
					}
					field.Set(reflect.ValueOf(res))

				case "StartTime", "EndTime":
					floatTime, err := strconv.ParseFloat(row[j], 64)
					if err != nil {
						log.Fatalf("Error parsing float string of %v: %v", e.Name, err)
					}
					timeClock, err := excelize.ExcelDateToTime(floatTime, false)
					if err != nil {
						log.Fatalf("Error converting float to clock of %v: %v", e.Name, err)
					}
					field.Set(reflect.ValueOf(timeClock))
				default:
					field.SetString(row[j])
				}
			} else {
				log.Printf("Error: Field %s not exist in struct", name)
			}
		}
		if e.StartTime.Hour() > e.EndTime.Hour() {
			newDate := e.EndTime.AddDate(0, 0, 1)
			metaValue.FieldByName("EndTime").Set(reflect.ValueOf(newDate))
		}
		employees = append(employees, *e)
		log.Printf("Added %v to collection %v", e.Name, e)
	}

	/*
		Fill spreadsheet.
		Here we fill spredsheet with calendar days, employees and schedule.
	*/
	var (
		sheetName    string = fmt.Sprintf("%v %v", month, year)
		sheetIndex          = f.NewSheet(sheetName)
		weekDaySlice        = []string{"", ""}
		weekDaysMap         = map[int]int{}
		monthDays    []int
		monthRow     = []any{"", "ФИО/Дата"}
	)
	log.Printf("Created sheet %v: %v\n", sheetIndex, sheetName)

	// Set row with calendar days
	for i := 1; i <= daysInMonth; i++ {
		wdi := int(time.Date(desiredTime.Year(), desiredTime.Month(), i, 0, 0, 0, 0, time.UTC).Weekday())
		weekDaysMap[i] = wdi
		weekDaySlice = append(weekDaySlice, daysOfWeek[weekDaysMap[i]])
	}
	monthDays = getKeys(weekDaysMap)
	sort.Ints(monthDays)

	for _, v := range monthDays {
		monthRow = append(monthRow, v)
	}
	monthRow = append(monthRow, "Норма часов, согласно производственному календарю", "Отработано в месяц (часов)", "Подпись работника")

	if err := f.SetSheetRow(sheetName, "A5", &monthRow); err != nil {
		log.Println("Sheet error C5", err)
	}
	if err := f.SetSheetRow(sheetName, "A6", &weekDaySlice); err != nil {
		log.Println("Sheet error C6", err)
	}

	//Set rows with employees
	i := 7
	for _, e := range employees {
		worktimeRow := []string{e.Job, e.Name}
		totalHoursRow := []any{}
		var totalHours time.Duration
		weekend := toInt(strings.Split(e.Weekend, ""))

		for l := 1; l < len(weekDaysMap)+1; l++ {
			isWeekend := false
			for _, v := range weekend {
				if weekDaysMap[l] == v {
					isWeekend = true
				}
			}
			if isWeekend == true {
				worktimeRow = append(worktimeRow, "B") // Pay attention to language
			} else {
				start := e.StartTime.Format("15:04")
				end := e.EndTime.Format("15:04")
				worktimeRow = append(worktimeRow, fmt.Sprintf("%v-%v", start, end))
			}
		}

		for _, v := range worktimeRow[2:] {
			switch v {
			case "B":
				totalHoursRow = append(totalHoursRow, "в")
			default:
				workDuration := e.EndTime.Sub(e.StartTime)
				totalHoursRow = append(totalHoursRow, workDuration.Hours())
				totalHours += workDuration
			}

		}
		totalHoursRow = append(totalHoursRow, "", totalHours.String())
		if err := f.SetSheetRow(sheetName, fmt.Sprintf("A%v", i), &worktimeRow); err != nil {
			log.Println("Sheet error worktimeRow", err)
		}
		if err := f.MergeCell(sheetName, fmt.Sprintf("A%v", i), fmt.Sprintf("A%v", i+1)); err != nil {
			log.Panic(err)
		}
		if err := f.MergeCell(sheetName, fmt.Sprintf("B%v", i), fmt.Sprintf("B%v", i+1)); err != nil {
			log.Panic(err)
		}
		if err := f.SetSheetRow(sheetName, fmt.Sprintf("C%v", i+1), &totalHoursRow); err != nil {
			log.Println("Sheet error totalHoursRow", err)
		}

		i += 2
	}
}
