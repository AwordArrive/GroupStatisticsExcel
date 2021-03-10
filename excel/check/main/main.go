package main

import (
	"fmt"
	"regexp"
	"strconv"

	excelize "github.com/360EntSecGroup-Skylar/excelize/v2"
)

type result struct {
	indexList []int

	data []string
}

var path string = "C:\\Users\\xsh\\Desktop\\T201.xlsx"

func main() {

	// filePath, _ := os.Getwd()
	// path = filePath + path
	rows := getCellsValue()
	re := checkFormatter(rows)
	writeResult(re)

}

func getCellsValue() [][]string {

	f, err := excelize.OpenFile(path)
	if err != nil {
		fmt.Println(err)
	}

	rows, err := f.GetRows("T201")
	return rows
}

func checkFormatter(dataList [][]string) []result {

	resultSlice := make([]result, 1000, 5000)
	for index, row := range dataList {

		if index < 7 {
			continue
		}
		indexList := make([]int, 1000)
		dataList := make([]string, 1000)

		var matchResult bool
		for index2, cell := range row {

			if cell == "" {
				continue
			}

			isMatch1, _ := regexp.MatchString(`\d*`, cell)
			isMatch2, _ := regexp.MatchString(`\d{4}-\d{2}-\d{2}`, cell)
			isMatch3, _ := regexp.MatchString(`\d\|^[\u4e00-\u9fa5]$`, cell)
			isMatch4, _ := regexp.MatchString(`^[\u4e00-\u9fa5]$`, cell)
			isMatch5, _ := regexp.MatchString(`[...]`, cell)

			matchResult = isMatch1 || isMatch2 || isMatch3 || isMatch4 || isMatch5

			if matchResult {
				indexList[index] = index2
				dataList[index] = cell
			}
		}

		result1 := result{indexList: indexList, data: dataList}
		resultSlice[index] = result1
	}

	return resultSlice
}

func writeResult(dataList []result) {

	f, err := excelize.OpenFile(path)
	if err != nil {
		fmt.Println(err)
	}
	sheetName := "sheet123"
	f.NewSheet(sheetName)

	for aaa, result := range dataList {

		if result.data == nil || len(result.data) == 0 {
			continue
		}

		indexs := result.indexList
		a := "A" + strconv.Itoa(aaa)
		f.SetSheetRow(sheetName, a, &result.data)

		for index := range indexs {

			rowNumerPre := int2Column(index)
			// fmt.Println(rowNumerPre)

			style, err := f.NewStyle(`{
				"font":
				{
					"bold": true,
					"italic": true,
					"family": "Times New Roman",
					"size": 36,
					"color": "#DC143C"
				}
			}`)
			if err != nil {
				fmt.Println(err)
			}
			rowNumber := rowNumerPre + strconv.Itoa(aaa)
			f.SetCellStyle(sheetName, rowNumber, rowNumber, style)
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
