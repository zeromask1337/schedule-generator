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
	Birthday    string
	StartTime   string
	EndTime     string
	Job         string
	PhoneNumber string
	Weekend     string
}

// Keys returns map keys
func Keys(m map[int]int) (keys []int) {
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
				field.SetString(row[j])
			} else {
				log.Printf("Error: Field %s not exist in struct", name)
			}
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
	log.Printf("Created sheet %v:%v\n", sheetIndex, sheetName)

	// Set row with calendar days
	for i := 1; i <= daysInMonth; i++ {
		wdi := int(time.Date(desiredTime.Year(), desiredTime.Month(), i, 0, 0, 0, 0, time.UTC).Weekday())
		weekDaysMap[i] = wdi
		weekDaySlice = append(weekDaySlice, daysOfWeek[weekDaysMap[i]])
	}
	monthDays = Keys(weekDaysMap)
	sort.Ints(monthDays)

	for _, v := range monthDays {
		monthRow = append(monthRow, v)
	}
	monthRow = append(monthRow, "норма часов, согласно производственному календарю", "отработано в месяц (часов)", "подпись работника")

	if err := f.SetSheetRow(sheetName, "A5", &monthRow); err != nil {
		log.Println("Sheet error C5", err)
	}
	if err := f.SetSheetRow(sheetName, "A6", &weekDaySlice); err != nil {
		log.Println("Sheet error C6", err)
	}

	//Set rows with employees
	i := 7
	for _, e := range employees {
		employeeWorkTime := []string{e.Job, e.Name}
		employeeTotalHoursRow := []any{}
		totalHours := 0
		employeeWeekend := toInt(strings.Split(e.Weekend, ""))
		for _, v := range weekDaysMap {
			switch v {
			case employeeWeekend[0], employeeWeekend[1]:
				employeeWorkTime = append(employeeWorkTime, "B") // Pay attention to language
			default:
				employeeWorkTime = append(employeeWorkTime, fmt.Sprintf("%v-%v", e.StartTime, e.EndTime))
			}
		}
		for _, v := range employeeWorkTime[2:] {
			switch v {
			case "B":
				employeeTotalHoursRow = append(employeeTotalHoursRow, "в")
			default:
				st, _ := strconv.Atoi(strings.ReplaceAll(e.StartTime, ":00", ""))
				et, _ := strconv.Atoi(strings.ReplaceAll(e.EndTime, ":00", ""))

				if et > st {
					employeeTotalHoursRow = append(employeeTotalHoursRow, et-st)
					totalHours += et - st
				} else {
					employeeTotalHoursRow = append(employeeTotalHoursRow, st-et)
					totalHours += st - et
				}

			}

		}
		employeeTotalHoursRow = append(employeeTotalHoursRow, "", totalHours)
		if err := f.SetSheetRow(sheetName, fmt.Sprintf("A%v", i), &employeeWorkTime); err != nil {
			log.Println("Sheet error employeeWorkTime", err)
		}
		f.MergeCell(sheetName, fmt.Sprintf("A%v", i), fmt.Sprintf("A%v", i+1))
		f.MergeCell(sheetName, fmt.Sprintf("B%v", i), fmt.Sprintf("B%v", i+1))
		if err := f.SetSheetRow(sheetName, fmt.Sprintf("C%v", i+1), &employeeTotalHoursRow); err != nil {
			log.Println("Sheet error employeeTotalHoursRow", err)
		}

		i += 2
	}
}
