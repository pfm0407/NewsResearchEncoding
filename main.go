package main

import (
	"fmt"
	"github.com/pfm0407/NewsResearchEncoding/internal/process"
	"github.com/xuri/excelize/v2"
	"strconv"
	"strings"
	"time"
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

	//f, err := excelize.OpenFile("input/" + "南海网我的.xlsx")
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

	/*
		cell, err := f.GetCellValue("Sheet1", "B2")
		if err != nil {
			fmt.Println(err)
			return
		}
		fmt.Println(cell)
	*/

	//Copy original sheet

	index, err := f.NewSheet("OriginalSheet")
	if err != nil {
		fmt.Println(err)
		return
	}
	err = f.CopySheet(1, index)

	var activeSheet string
	var activeCol string
	var activeCell string

	activeSheet = "Sheet1"
	titleCol := 1
	contentCol := 2

	rows, err := f.GetRows(activeSheet)
	if err != nil {
		fmt.Println(err)
		return
	}

	//SetRelatedMark
	//Todo

	//SetOriginalMark

	activeCol = "Q"
	for i, row := range rows {
		activeCell = activeCol + strconv.Itoa(i+1)
		switch {
		case strings.Contains(row[contentCol], "新华社"):
			err = f.SetCellDefault(activeSheet, activeCell, "Non-original") //Non-original
		case strings.Contains(row[contentCol], "本报记者"):
			err = f.SetCellDefault(activeSheet, activeCell, "Original") //Original
		default:
			err = f.SetCellDefault(activeSheet, activeCell, "Unknown") //default:unknown
		}

		//fmt.Print(row[2], "\t")
		if err != nil {
			fmt.Println(err)
			return
		}
	}

	//SetDeleteMark

	//SetTheme

	activeCol = "R"

	//“政治、经济、文化、社会、生态、科技、法治、教育、军事、其他”

	//一般规则
	basicTitleKeywords := [][]string{
		{"政治", "政治", "两国关系"},
		{"经济", "经济", "外贸", "投资", "自贸", "营商"},
		{"文化", "文化", "人文"},
		{"社会", "社会", "旅游推介"},
		{"生态", "生态", "气候", "绿色产业", "低碳"},
		{"科技", "科技"},
		{"法治", "法治", "贪腐", "腐败", "立法", "司法", "公正"},
		{"教育", "教育", "学校", "中学", "小学", "大学", "教师"},
		{"军事", "军事"},
	}

	basicContentKeywords := [][]string{
		//{"政治", "政治", "两国关系"},
		{"经济", "数字人民币"},
		//{"文化", "文化", "人文"},
		{"社会", "文旅局", "旅文局", "文旅推荐官"},
		{"生态", "生态环境部", "绿色可持续发展"},
		//{"科技", "科技"},
		//{"法治", "法治", "贪腐", "腐败", "立法", "司法", "公正"},
		{"教育", "上学"},
		//{"军事", "军事"},
	}

	for i, row := range rows {
		activeCell = activeCol + strconv.Itoa(i+1)

		for j := 0; j < len(basicTitleKeywords); j++ {
			for k := 1; k < len(basicTitleKeywords[j]); k++ {
				if strings.Contains(row[titleCol], basicTitleKeywords[j][k]) {
					err = f.SetCellDefault(activeSheet, activeCell, basicTitleKeywords[j][0])
				}
			}
		}

		for j := 0; j < len(basicContentKeywords); j++ {
			for k := 1; k < len(basicContentKeywords[j]); k++ {
				if strings.Contains(row[contentCol], basicContentKeywords[j][k]) {
					err = f.SetCellDefault(activeSheet, activeCell, basicContentKeywords[j][0])
				}
			}
		}

		//fmt.Print(row[2], "\t")
		if err != nil {
			fmt.Println(err)
			return
		}
	}

	//特殊规则
	/*
		for i, row := range rows {
			activeCell := activeCol + strconv.Itoa(i+1)
			switch {
			case strings.Contains(row[themeCol], "外贸"):
				err = f.SetCellDefault(activeSheet, activeCell, "经济")
			case strings.Contains(row[themeCol], "自贸"):
				err = f.SetCellDefault(activeSheet, activeCell, "经济")
			case strings.Contains(row[themeCol], "经济"):
				err = f.SetCellDefault(activeSheet, activeCell, "经济")
			default:
				err = f.SetCellDefault(activeSheet, activeCell, "Unknown") //default:unknown
			}

			//fmt.Print(row[2], "\t")
			if err != nil {
				fmt.Println(err)
				return
			}
		}
	*/

	activeCol = "S"

	countUnique := 1
	var countDuplicate int
	for i, row1 := range rows {
		countDuplicate = 2
		activeCellA := activeCol + strconv.Itoa(i+1)
		contentCellA, _ := f.GetCellValue(activeSheet, activeCellA)
		if contentCellA != "" {
			continue
		}
		for j, row2 := range rows {
			if i == j {
				continue
			}
			//fmt.Println(row2[titleCol] + "_" + row1[titleCol])
			if strings.Contains(row2[titleCol], row1[titleCol]) {
				//fmt.Println("FindDuplicate")
				activeCellB := activeCol + strconv.Itoa(j+1)
				err = f.SetCellDefault(activeSheet, activeCellA, strconv.Itoa(countUnique)+"_"+"1")
				err = f.SetCellDefault(activeSheet, activeCellB, strconv.Itoa(countUnique)+"_"+strconv.Itoa(countDuplicate))
				countDuplicate += 1
			}
		}
		if countDuplicate != 2 {
			countUnique += 1
		}
	}

	//SetCity&Area
	//SetImportance

	//ExecuteDelete

	newName := "output/" + "Result" + time.Now().Format("20060102_150405") + ".xlsx"
	err = f.SaveAs(newName)
	fmt.Println("Saving...")
	if err != nil {
		fmt.Println(err)
		return
	}

}
