package main

import (
	"fmt"
	"github.com/xuri/excelize/v2"
	"io/ioutil"
	"os"
	"path"
	"strconv"
)

type Man struct {
	IdCard string
	Value  float64
	Items  []*Man
}

var (
	outPath     = "./output"
	outputSheet = "Sheet1"
	manA        []*Man
	manB        []*Man
)

func main() {
	files, err := ioutil.ReadDir(".")
	if err != nil {
		println(err)
		return
	}

	err = os.MkdirAll(outPath, os.ModePerm)
	if err != nil {
		println(err)
		return
	}
	for _, f := range files {
		ext := path.Ext(f.Name())
		if ext == ".xlsx" {
			fmt.Println(path.Base(f.Name()))
			handle(f.Name())
		}
	}

}

func handle(fileName string) {
	outputFileName := outPath + "/" + fileName

	manA = make([]*Man, 0, 10)
	manB = make([]*Man, 0, 10)
	read(fileName)
	match()
	write(outputFileName)

}

func read(fileName string) {

	f, err := excelize.OpenFile(fileName)
	defer f.Close()
	if err != nil {
		fmt.Println(err)
		return
	}

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
}

func match() {

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

}

func write(outputFileName string) {

	nf := excelize.NewFile()
	defer nf.Close()
	style, err := nf.NewStyle(&excelize.Style{
		Fill: excelize.Fill{Type: "pattern", Color: []string{"#fff000"}, Pattern: 1},
	})
	index := nf.NewSheet(outputSheet)
	row := 0
	for _, man := range manA {
		row = row + 1
		_ = nf.SetCellValue(outputSheet, axis(row, 1), man.IdCard)
		_ = nf.SetCellValue(outputSheet, axis(row, 2), man.Value)
		err = nf.SetCellStyle(outputSheet, axis(row, 1), axis(row, 2), style)

		if man.Items != nil {
			for _, item := range man.Items {
				row = row + 1
				_ = nf.SetCellValue(outputSheet, axis(row, 1), item.IdCard)
				_ = nf.SetCellValue(outputSheet, axis(row, 2), item.Value)
			}
		}
	}
	_ = nf.SetColWidth(outputSheet, "A", "B", 20)
	if err != nil {
		println(err)
		return
	}

	// 设置工作簿的默认工作表
	nf.SetActiveSheet(index)
	// 根据指定路径保存文件
	if err := nf.SaveAs(outputFileName); err != nil {
		fmt.Println(err)
	}

}
func axis(row, col int) string {
	colN, _ := excelize.ColumnNumberToName(col)
	return colN + strconv.Itoa(row)
}
