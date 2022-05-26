package main

import (
	"fmt"
	"github.com/xuri/excelize/v2"
	"io/ioutil"
	"log"
	"os"
	"path"
	"strconv"
)

type Man struct {
	IdCard string
	Value  float64
	Items  []*Man
	Sum    float64
}

var (
	outPath = "./output"
	manA    []*Man
	manB    []*Man

	mapA map[string]*Man
	mapB map[string]*Man
)

func main() {
	log.Println("开始执行程序...")
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
			log.Printf("读取文件: %s\n", path.Base(f.Name()))
			handle(f.Name())
		}
	}

	askExit()

}

func clearCache() {
	manA = nil
	manB = nil
	mapA = nil
	mapB = nil
}

func handle(fileName string) {
	outputFileName := outPath + "/" + "_" + fileName
	manA = make([]*Man, 0, 10)
	manB = make([]*Man, 0, 10)
	mapA = make(map[string]*Man)
	mapB = make(map[string]*Man)
	read(fileName)
	match2()
	write(outputFileName)
	clearCache()

}

func read(fileName string) {

	f, err := excelize.OpenFile(fileName)
	f.Close()
	if err != nil {
		log.Println(err)
		return
	}

	rows, err := f.GetRows("Sheet1")
	for _, row := range rows {
		if len(row) >= 2 {
			k := row[0]
			if k != "" {
				v, _ := strconv.ParseFloat(row[1], 64)
				data := &Man{
					IdCard: k,
					Value:  v,
				}
				manA = append(manA, data)
				mapA[k] = data
			}
		}

	}

	rows, err = f.GetRows("Sheet2")
	for _, row := range rows {
		if len(row) >= 2 {
			k := row[0]
			if k != "" {
				v, _ := strconv.ParseFloat(row[1], 64)
				data := &Man{
					IdCard: k,
					Value:  v,
				}
				manB = append(manB, data)
				mapB[k] = data
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

func match2() {
	i := -1
	for _, a := range manA {

		//首先要查询b组里面是否有自己
		self := mapB[a.IdCard]
		if self != nil {
			a.Items = append(a.Items, self)
			a.Sum = a.Sum + self.Value
		}

		for a.Sum < a.Value {
			i++
			if i >= len(manB) {
				break
			}

			b := manB[i]
			//在上面会首先把自己添加进来，这里忽略自己
			if b.IdCard == a.IdCard {
				continue
			}

			//如果a队伍中存在b，跳过不处理
			if mapA[b.IdCard] != nil {
				continue
			}

			a.Items = append(a.Items, b)
			a.Sum = a.Sum + b.Value
		}
	}
}

func write(outputFileName string) {
	nf := excelize.NewFile()
	defer nf.Close()
	styleA, err := nf.NewStyle(&excelize.Style{
		Fill: excelize.Fill{Type: "pattern", Color: []string{"#FFF000"}, Pattern: 1},
	})
	styleB, err := nf.NewStyle(&excelize.Style{
		Fill: excelize.Fill{Type: "pattern", Color: []string{"#008B45"}, Pattern: 1},
	})

	sheet1 := "Sheet1"
	for i, man := range manA {
		row := i + 1
		colA := axis(row, 1)
		colB := axis(row, 2)
		err = nf.SetCellValue(sheet1, colA, man.IdCard)
		err = nf.SetCellValue(sheet1, colB, man.Value)
		err = nf.SetCellStyle(sheet1, colA, colB, styleA)
	}

	sheet2 := "Sheet2"
	index := nf.NewSheet(sheet2)
	for i, man := range manB {
		row := i + 1
		colA := axis(row, 1)
		colB := axis(row, 2)
		err = nf.SetCellValue(sheet2, colA, man.IdCard)
		err = nf.SetCellValue(sheet2, colB, man.Value)
		err = nf.SetCellStyle(sheet2, colA, colB, styleB)
	}

	sheet3 := "Sheet3"
	index = nf.NewSheet(sheet3)
	row := 0
	for _, man := range manA {
		row = row + 1
		colA := axis(row, 1)
		colB := axis(row, 2)
		err = nf.SetCellValue(sheet3, colA, man.IdCard)
		err = nf.SetCellValue(sheet3, colB, man.Value)
		err = nf.SetCellStyle(sheet3, colA, colB, styleA)
		if man.Items != nil {
			for _, item := range man.Items {
				row = row + 1
				colAA := axis(row, 1)
				colBB := axis(row, 2)
				err = nf.SetCellValue(sheet3, colAA, item.IdCard)
				err = nf.SetCellValue(sheet3, colBB, item.Value)
				err = nf.SetCellStyle(sheet3, colAA, colBB, styleB)
			}
		}
	}
	_ = nf.SetColWidth(sheet1, "A", "B", 20)
	_ = nf.SetColWidth(sheet2, "A", "B", 20)
	_ = nf.SetColWidth(sheet3, "A", "B", 20)
	if err != nil {
		println(err)
		return
	}

	// 设置工作簿的默认工作表
	nf.SetActiveSheet(index)
	// 根据指定路径保存文件
	if err := nf.SaveAs(outputFileName); err != nil {
		log.Println(err)
	}
	log.Printf("写入文件%s\t", outputFileName)

}
func axis(row, col int) string {
	colN, _ := excelize.ColumnNumberToName(col)
	return colN + strconv.Itoa(row)
}

func init() {
	settingLog()
}

func settingLog() {
	//f, err := os.OpenFile("log", os.O_CREATE|os.O_APPEND|os.O_RDWR, os.ModePerm)
	//if err != nil {
	//	return
	//}
	//log.SetOutput(f)
}

func askExit() {
	fmt.Printf("执行完成按任意键退出...")
	b := make([]byte, 1)
	os.Stdin.Read(b)
}
