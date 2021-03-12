package main

import (
	"fmt"
	"os"
	"strconv"
	"strings"

	excelize "github.com/360EntSecGroup-Skylar/excelize/v2"
)

type result struct {
	number int

	indexSlice []int

	dataSlice []string
}

// var path string = "C:\\Users\\xsh\\Desktop\\t.xlsx"

var path string = "\\ttt.xlsx"

func main() {

	filePath, _ := os.Getwd()
	path = filePath + path
	rows := getCellsValue()
	var resultSlice []result
	for i, row := range rows {

		if i < 3 {
			continue
		}
		indexSlice := checkFormatter(row)
		if indexSlice == nil || len(indexSlice) == 0 {
			continue
		}
		result := result{i, indexSlice, row}
		resultSlice = append(resultSlice, result)
	}

	writeResult(resultSlice)

	// 重新写入以解决格式不生效的问题
	writeResult(resultSlice)
}

func getCellsValue() [][]string {

	f, err := excelize.OpenFile(path)
	if err != nil {
		fmt.Println(err)
	}

	rows, err := f.GetRows("t")
	return rows
}

func checkFormatter(dataList []string) []int {

	var indexSlice []int
	for index, cell := range dataList {

		if cell == "" {
			continue
		}
		matchResult := strings.Contains(cell, " ")
		if matchResult {
			indexSlice = append(indexSlice, index)
		}
	}

	return indexSlice
}

func writeResult(resultSlice []result) {

	f, err := excelize.OpenFile(path)
	if err != nil {
		fmt.Println(err)
	}
	sheetName := "sheet123"
	f.NewSheet(sheetName)
	for rowNum, result := range resultSlice {

		rear := append([]string{}, result.dataSlice[1:]...)
		result.dataSlice = append(result.dataSlice[0:0], strconv.Itoa(result.number))
		result.dataSlice = append(result.dataSlice, rear...)

		rowNum = rowNum + 1
		f.SetSheetRow(sheetName, "A"+strconv.Itoa(rowNum), &result.dataSlice)
		for _, index := range result.indexSlice {

			style, err := f.NewStyle(`{
				"font":
					{
						"color": "#DC143C"
					}
				}`)
			if err != nil {
				fmt.Println(err)
			}

			cellNumber := int2Column(index) + strconv.Itoa(rowNum)
			f.SetCellStyle(sheetName, cellNumber, cellNumber, style)
		}
	}
	f.Save()
}

func int2Column(column int) string {

	var digArr [100]string
	mod := column % 26
	col := (column - 1) / 26
	for i := 0; i < col; i++ {
		digArr[i] = "A"
	}
	digArr[col] = dig2Char(mod)

	result := ""
	for _, i := range digArr { //遍历数组中所有元素追加成string
		result += i
	}
	return result
}

func dig2Char(dig int) string {

	chars := []string{"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"}

	return chars[dig]
}
