package main

import (
	"container/list"
	"encoding/json"
	"fmt"
	"os"
	"sort"
	"strconv"
	"strings"

	excelize "github.com/360EntSecGroup-Skylar/excelize/v2"
)

var path string = "\\问题记录表-河南-20210124.xlsx"

var startRowIndex int

func main() {

	filePath, _ := os.Getwd()
	path = filePath + path

	cellsValue := getCellsValue()
	maps := getCity()

	resultMap := make(map[string]int64, 32)
	for e := cellsValue.Front(); e != nil; e = e.Next() {

		cellValue := fmt.Sprintf("%v", e.Value)
		if cellValue != "" {
			city := converToCity(cellValue, maps)
			num := resultMap[city]
			resultMap[city] = num + 1
		}
	}

	fmt.Println(resultMap)
	writeResult(resultMap)
}

func getCellsValue() *list.List {

	f, err := excelize.OpenFile(path)
	if err != nil {
		fmt.Println(err)
	}

	rows, err := f.GetRows("XX省")
	if err != nil {
		fmt.Println(err)
		return nil
	}

	fmt.Println("请输入统计开始行号：")
	fmt.Scanf("%d", &startRowIndex) //注意使用%s读取输入字符串只能读取到空白符之前

	cityList := list.New()
	for index, row := range rows {

		if row == nil {
			continue
		}

		if index >= startRowIndex-1 {
			cityList.PushFront(row[3])
		}
	}

	return cityList
}

func getCity() map[string][]string {

	s := "[{\"郑州\":[\"中原\",\"二七\",\"管城回族\",\"金水\",\"上街\",\"惠济\",\"中牟\",\"经济技术开发\",\"高新技术产业开发\",\"航空港经济综合实验\",\"巩义\",\"荥阳\",\"新密\",\"新郑\",\"登封\"]},{\"开封\":[\"龙亭\",\"顺河回族\",\"鼓楼\",\"禹王台\",\"祥符\",\"杞\",\"通许\",\"尉氏\",\"兰考\"]},{\"洛阳\":[\"老城\",\"西工\",\"瀍河\",\"涧西\",\"吉利\",\"洛龙\",\"孟津\",\"新安\",\"栾川\",\"嵩\",\"汝阳\",\"宜阳\",\"洛宁\",\"伊川\",\"高新技术产业开发\",\"偃师\"]},{\"平顶山\":[\"新华\",\"卫东\",\"石龙\",\"湛河\",\"宝丰\",\"叶\",\"鲁山\",\"郏\",\"高新技术产业开发\",\"新城\",\"舞钢\",\"汝州\"]},{\"安阳\":[\"文峰\",\"北关\",\"殷都\",\"龙安\",\"安阳\",\"汤阴\",\"滑\",\"内黄\",\"高新技术产业开发\",\"林州\"]},{\"鹤壁\":[\"鹤山\",\"山城\",\"淇滨\",\"浚\",\"淇\",\"经济技术开发\"]},{\"新乡\":[\"红旗\",\"卫滨\",\"凤泉\",\"牧野\",\"新乡\",\"获嘉\",\"原阳\",\"延津\",\"封丘\",\"高新技术产业开发\",\"经济技术开发\",\"平原城乡一体化示范\",\"卫辉\",\"辉\",\"长垣\"]},{\"焦作\":[\"解放\",\"中站\",\"马村\",\"山阳\",\"修武\",\"博爱\",\"武陟\",\"温\",\"城乡一体化示范\",\"沁阳\",\"孟州\"]},{\"濮阳\":[\"华龙\",\"清丰\",\"南乐\",\"范\",\"台前\",\"濮阳\",\"工业园\",\"经济技术开发\"]},{\"许昌\":[\"魏都\",\"建安\",\"鄢陵\",\"襄城\",\"经济开发\",\"禹州\",\"长葛\"]},{\"漯河\":[\"源汇\",\"郾城\",\"召陵\",\"舞阳\",\"临颍\",\"经济技术开发\"]},{\"三门峡\":[\"湖滨\",\"陕州\",\"渑池\",\"卢氏\",\"经济开发\",\"义马\",\"灵宝\"]},{\"南阳\":[\"宛城\",\"卧龙\",\"南召\",\"方城\",\"西峡\",\"镇平\",\"内乡\",\"淅川\",\"社旗\",\"唐河\",\"新野\",\"桐柏\",\"高新技术产业开发\",\"城乡一体化示范\",\"邓州\"]},{\"商丘\":[\"梁园\",\"睢阳\",\"民权\",\"睢\",\"宁陵\",\"柘城\",\"虞城\",\"夏邑\",\"豫东综合物流产业聚集\",\"商丘经济开发\",\"永城\"]},{\"信阳\":[\"浉河\",\"平桥\",\"罗山\",\"光山\",\"新\",\"商城\",\"固始\",\"潢川\",\"淮滨\",\"息\",\"高新技术产业开发\"]},{\"周口\":[\"川汇\",\"淮阳\",\"扶沟\",\"西华\",\"商水\",\"沈丘\",\"郸城\",\"太康\",\"鹿邑\",\"经济开发\",\"项城\"]},{\"驻马店\":[\"驿城\",\"西平\",\"上蔡\",\"平舆\",\"正阳\",\"确山\",\"泌阳\",\"汝南\",\"遂平\",\"新蔡\",\"经济开发\"]},{\"济源\":[\"沁园街道\",\"济水街道\",\"北海街道\",\"天坛街道\",\"玉泉街道\",\"克井镇\",\"五龙口镇\",\"轵城镇\",\"承留镇\",\"邵原镇\",\"坡头镇\",\"梨林镇\",\"大峪镇\",\"思礼镇\",\"王屋镇\",\"下冶镇\"]}]"

	maps := make(map[string][]string, 64)
	var dats []map[string][]string
	if err := json.Unmarshal([]byte(s), &dats); err == nil {
		for _, m := range dats {
			for k, v := range m {
				maps[k] = v
			}
		}
	} else {
		fmt.Println(err)
	}
	return maps
}

func converToCity(cellValue string, maps map[string][]string) string {

	if strings.Contains(cellValue, "JJohn") {
		return "其他"
	}
	if strings.Contains(cellValue, "linda 快乐") {
		return "其他"
	}
	if strings.Contains(cellValue, "蝴蝶") {
		return "其他"
	}
	if strings.Contains(cellValue, "栀子花开") {
		return "其他"
	}
	if strings.Contains(cellValue, "简单") {
		return "其他"
	}
	if strings.Contains(cellValue, "Vera Shi") || strings.Contains(cellValue, "vera Shi") {
		return "其他"
	}

	if strings.Contains(cellValue, "东") {
		return "其他"
	}
	if strings.Contains(cellValue, "东区-张威") {
		return "其他"
	}

	if strings.Contains(cellValue, "张") {
		return "商丘"
	}
	if strings.Contains(cellValue, "民政局") {
		return "孟州市"
	}
	if strings.Contains(cellValue, "果儿") {
		return "洛阳"
	}
	if strings.Contains(cellValue, "垚窕") {
		return "安阳"
	}
	if strings.Contains(cellValue, "幸福像花一样红") {
		return "其他"
	}
	if strings.Contains(cellValue, "吉祥如意") {
		return "洛阳"
	}
	if strings.Contains(cellValue, "省厅") {
		return "省厅"
	}
	if strings.Contains(cellValue, "fengfeng") {
		return "濮阳"
	}
	if strings.Contains(cellValue, "努力奋斗") {
		return "新乡"
	}
	for k, v := range maps {

		if strings.Contains(cellValue, k) {
			return k
		}
		for _, contry := range v {

			if strings.Contains(cellValue, contry) {
				return k
			}
		}
	}
	return "其他"
}

func writeResult(maps map[string]int64) {

	f, err := excelize.OpenFile(path)
	if err != nil {
		fmt.Println(err)
	}
	sheetName := fmt.Sprintf("%d", startRowIndex)
	f.NewSheet(sheetName)
	f.SetCellValue(sheetName, "A1", "地区")
	f.SetCellValue(sheetName, "B1", "数量")

	type kv struct {
		Key   string
		Value int64
	}
	var ss []kv
	for k, v := range maps {
		ss = append(ss, kv{k, v})
	}
	sort.Slice(ss, func(i, j int) bool {
		return ss[i].Value > ss[j].Value // 降序
	})

	var index int = 1
	for _, kv := range ss {
		index = index + 1
		fmt.Printf("%s, %d\n", kv.Key, kv.Value)
		a := "A" + strconv.Itoa(index)
		b := "B" + strconv.Itoa(index)
		f.SetCellValue(sheetName, a, kv.Key)
		f.SetCellValue(sheetName, b, kv.Value)
	}

	var format string = `{
        "type": "bar",
        "series": [
        {
            "name": "{{sheetName}}!$A$2",
            "categories": "{{sheetName}}!$A$2:$A${{dataLength}}",
            "values": "{{sheetName}}!$B$2:$B${{dataLength}}"
        }],
        "format":
        {
            "x_scale": 1.0,
            "y_scale": 1.0,
            "x_offset": 15,
            "y_offset": 10,
            "print_obj": true,
            "lock_aspect_ratio": false,
            "locked": false
        },
        "legend":
        {
			"none": true,
            "position": "left",
            "show_legend_key": false
        },
        "title":
        {
            "name": "统计系统问题统计"
        },
        "plotarea":
        {
            "show_bubble_size": true,
            "show_cat_name": false,
            "show_leader_lines": true,
            "show_percent": true,
            "show_series_name": false,
            "show_val": true
        },
        "show_blanks_as": "zero"
    }`

	format = strings.Replace(format, "{{sheetName}}", sheetName, 3)
	format = strings.Replace(format, "{{dataLength}}", fmt.Sprintf("%d", index), 2)

	if err := f.AddChart(sheetName, "E1", format); err != nil {
		fmt.Println(err)
	}

	f.Save()
}
