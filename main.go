package main

import (
	"fmt"
	"github.com/pfm0407/NewsResearchEncoding/internal/process"
)

func OriginalMark() int32 {
	//Start
	fmt.Println("OriginalMark_Start")
	var optRowNum int32

	//Done
	fmt.Println("OriginalMark_Done")
	return optRowNum
}

func main() {

	process.ThemeSet()

	/*
		f, err := excelize.OpenFile("input/" + "test1.xlsx")
		if err != nil {
			fmt.Println(err)
			return
		}
		defer func() {
			// Close the spreadsheet.
			if err := f.Close(); err != nil {
				fmt.Println(err)
			}
		}()
		// Get value from cell by given worksheet name and cell reference.

			cell, err := f.GetCellValue("Sheet1", "B2")
			if err != nil {
				fmt.Println(err)
				return
			}
			fmt.Println(cell)

		//Copy original sheet

			index, err := f.NewSheet("OriginalSheet")
			if err != nil {
				fmt.Println(err)
				return
			}
			err = f.CopySheet(1, index)


		rows, err := f.GetRows("Sheet1")
		if err != nil {
			fmt.Println(err)
			return
		}

		//SetRelatedMark
		//Todo

		//SetOriginalMark
		for i, row := range rows {
			switch {
			case strings.Contains(row[2], "新华社"):
				err = f.SetCellDefault("Sheet1", "Q"+strconv.Itoa(i+1), "Non-original") //Non-original
			case strings.Contains(row[2], "本报记者"):
				err = f.SetCellDefault("Sheet1", "Q"+strconv.Itoa(i+1), "Original") //Original
			default:
				err = f.SetCellDefault("Sheet1", "Q"+strconv.Itoa(i+1), "Unknown") //default:unknown
			}

			//fmt.Print(row[2], "\t")
			if err != nil {
				fmt.Println(err)
				return
			}
		}

		//SetDeleteMark

		//SetTheme
		for i, row := range rows {
			switch {
			case strings.Contains(row[1], "外贸"):
				err = f.SetCellDefault("Sheet1", "R"+strconv.Itoa(i+1), "经济")
			case strings.Contains(row[2], "本报记者"):
				err = f.SetCellDefault("Sheet1", "R"+strconv.Itoa(i+1), "经济")
			default:
				err = f.SetCellDefault("Sheet1", "R"+strconv.Itoa(i+1), "Unknown") //default:unknown
			}

			//fmt.Print(row[2], "\t")
			if err != nil {
				fmt.Println(err)
				return
			}
		}
		//SetCity&Area
		//SetImportance

		//ExecuteDelete

		err = f.Save()
		fmt.Println("Saving...")
		if err != nil {
			fmt.Println(err)
			return
		}

	*/
}
