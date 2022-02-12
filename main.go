package main

import (
	"fmt"
	"log"
	"os"
	"strconv"

	"github.com/xuri/excelize/v2"
)

func main() {

	userInput := os.Args
	if len(userInput) < 2 {
		fmt.Println("Please specify xlsx file as input")
		os.Exit(1)
	}
	fmt.Println(userInput)
	f, err := excelize.OpenFile(userInput[1])

	if err != nil {
		log.Fatal(err)
	}

	rows, err := f.GetRows("Sheet1")
	if err != nil {
		log.Fatal(err)
	}
	fmt.Println(len(rows))

	err1 := f.InsertCol("Sheet1", "H")
	if err1 != nil {
		log.Fatal(err1)
	}
	err2 := f.InsertCol("Sheet1", "I")
	if err2 != nil {
		log.Fatal(err1)
	}
	f.SetCellValue("Sheet1", "H1", "REMARKS")

	for i, row := range rows {
		for j, col := range row {

			if j == 5 && col == "FORECAST" {
				continue

			} else if j == 5 && i > 0 {
				forecast, err := strconv.ParseFloat(col, 64)
				if err != nil {
					log.Fatal(err)
				}
				if forecast <= 1 {
					columnH := fmt.Sprintf("H%d", i+1)
					columnG := fmt.Sprintf("G%d", i+1)
					//fmt.Println(columnH)
					val, err := f.GetCellValue("Sheet1", columnG)
					if err != nil {
						log.Fatal(err)
					}
					//fmt.Println(val)
					sbo, err := strconv.ParseFloat(val, 64)
					if err != nil {
						log.Fatal(err)
					}
					if sbo > 0 {
						f.SetCellValue("Sheet1", columnH, "SBO")
					} else {

						f.SetCellValue("Sheet1", columnH, "LOWFC")
					}

				} else {

					columnH := fmt.Sprintf("H%d", i+1)
					columnG := fmt.Sprintf("G%d", i+1)
					//fmt.Println(columnH)
					val, err := f.GetCellValue("Sheet1", columnG)
					if err != nil {
						log.Fatal(err)
					}
					//fmt.Println(val)
					sbo, err := strconv.ParseFloat(val, 64)
					if err != nil {
						log.Fatal(err)
					}
					if sbo > 0 {
						f.SetCellValue("Sheet1", columnH, "SBO")
					} else {

						f.SetCellValue("Sheet1", columnH, "HS")
					}

				}

			}

		}
	}
	f.SaveAs("Output.xlsx")
	f.Close()
}
