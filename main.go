package main

import (
	"fmt"
	"github.com/xuri/excelize/v2"
	"strconv"
)

type Man struct {
	IdCard string
	Value  float64
	Items  []*Man
}

func main() {
	f, err := excelize.OpenFile("table.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}

	manA := make([]*Man, 0, 10)
	rows, err := f.GetRows("Sheet1")
	for _, row := range rows {
		if len(row) >= 2 {
			k1 := row[0]
			if k1 != "" {
				v1, _ := strconv.ParseFloat(row[1], 64)
				manA = append(manA, &Man{
					IdCard: k1,
					Value:  v1,
				})
			}
		}

	}

	manB := make([]*Man, 0, 10)
	rows, err = f.GetRows("Sheet2")
	for _, row := range rows {
		if len(row) >= 2 {
			k1 := row[0]
			if k1 != "" {
				v1, _ := strconv.ParseFloat(row[1], 64)
				manB = append(manB, &Man{
					IdCard: k1,
					Value:  v1,
				})
			}
		}

	}

	sumVal := 0.0
	i := 0
	for _, b := range manB {
		if i >= len(manA) {
			break
		}
		a := manA[i]
		a.Items = append(a.Items, b)
		sumVal = sumVal + b.Value
		if sumVal >= a.Value {
			i = i + 1
			sumVal = 0.0
		}
	}

	//style, err := f.NewStyle(`{"fill":{"type":"pattern","color":["#E0EBF5"],"pattern":1}}`)
	outputSheet := "Sheet1"
	style, err := f.NewStyle(&excelize.Style{
		Fill: excelize.Fill{Type: "pattern", Color: []string{"#fff000"}, Pattern: 1},
	})
	index := f.NewSheet(outputSheet)
	row := 0
	for _, man := range manA {
		row = row + 1
		f.SetCellValue(outputSheet, axis(row, 1), man.IdCard)
		f.SetCellValue(outputSheet, axis(row, 2), man.Value)
		err = f.SetCellStyle(outputSheet, axis(row, 1), axis(row, 2), style)

		if man.Items != nil {
			for _, item := range man.Items {
				row = row + 1
				f.SetCellValue(outputSheet, axis(row, 1), item.IdCard)
				f.SetCellValue(outputSheet, axis(row, 2), item.Value)
			}
		}
	}
	f.SetColWidth(outputSheet, "A", "B", 20)
	if err != nil {

		println(err)
	}

	// 设置工作簿的默认工作表
	f.SetActiveSheet(index)
	// 根据指定路径保存文件
	if err := f.SaveAs("output.xlsx"); err != nil {
		fmt.Println(err)
	}

}
func axis(row, col int) string {
	colN, _ := excelize.ColumnNumberToName(col)
	return colN + strconv.Itoa(row)
}
