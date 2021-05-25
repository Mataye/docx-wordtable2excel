package main

import (
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize"
	"io/ioutil"
	"strings"
	"time"

)

type Object map[string]interface{}

func (o Object) Set(key string,val interface{})  {
	if o == nil {
		o = Object{}
	}
	o[key]= val
}

// 横列展示的表格
var FocusField2ExcelMap  = map[string]string{
	// TODO
	// 这里填需要映射的字段
}

var FocusField2ExcelArr = []string{
	//	TODO
	//  填写	FocusField2ExcelMap 中的 key
	//	因为 map 的无序性
}

func main() {
	files, _ := ioutil.ReadDir("./")
	fileDataArr := make([]Object, 0)
	for _, f := range files {
		//fmt.Println(f.Name())
		if f.IsDir() {
			continue
		}
		fileName := f.Name()
		fileSplits := strings.Split(fileName, ".")
		if len(fileSplits) != 2 {
			continue
		}

		if fileSplits[1] == "docx" {
			docTable, err := openDocxFile(fileName)
			if nil != err {
				fmt.Printf("打开文件 %s 失败 \n", fileName)
				continue
			}

			if len(docTable) == 0 {
				fmt.Printf("读取文件 %s 失败，没有可使用的数据 \n", fileName)
				continue
			}
			dataMap := pickColumnTableFiled(docTable)
			dataMap.Set("原始文件名", fileName)
			fileDataArr = append(fileDataArr, dataMap)
		} else {
			fmt.Printf("暂不支持 %s 格式的文件 \n", fileSplits[1])
		}
	}

	dateStr := time.Now().Format("20060102")
	excelName := fmt.Sprintf("%s_接诉即办工单.xlsx", dateStr)

	err := Save2Excel(excelName, fileDataArr)
	if err != nil {
		fmt.Printf("生成 excel 文件失败 \n")
	}
	fmt.Printf("生成 excel 文件 %s 成功 \n", excelName)
}

func openDocxFile(filename string) ([][]string, error) {
	reader, readerFiles :=  UnpackDocx(filename)
	if reader == nil || len(readerFiles) == 0 {
		return nil,fmt.Errorf("empty docx reader")
	}
	defer reader.Close()
	docFile := RetrieveWordDoc(readerFiles)
	doc := OpenWordDoc(*docFile)
	content := WordDocToString(doc)
	exectContent :=Extract(content)
	t := exectContent.Body.Table.TableRow

	var table [][]string
	for _, r := range t {
		var row []string
		for _, c := range r.TableColumn {
			row = append(row, strings.Join(c.Cell, ""))
		}
		table = append(table, row)
	}

	return table,nil
}

func pickColumnTableFiled(table [][]string) Object {
	dataMap := Object{}
	for _, row := range table {
		for index, field := range row {
			if newField, ok := FocusField2ExcelMap[field]; ok {
				if index+1 <= len(row) {
					dataMap.Set(newField, row[index+1])
				}
			}
		}
	}
	return dataMap
}

func Save2Excel(excelName string,fileDataArr []Object) error {
	f :=  excelize.NewFile()
	index := f.NewSheet("Sheet1")

	startColumnWord := "A"
	startRowIndex := 1
	for offset,field := range FocusField2ExcelArr {
		tmpColumnWord := transWord(startColumnWord, int32(offset))
		column := fmt.Sprintf("%s%d", tmpColumnWord, startRowIndex)

		f.SetCellValue("Sheet1", column, field)
	}

	startRowIndex++
	for _,obj :=range fileDataArr {
		for offset, field := range FocusField2ExcelArr {
			tmpColumnWord := transWord(startColumnWord, int32(offset))
			column := fmt.Sprintf("%s%d", tmpColumnWord, startRowIndex)
			f.SetCellValue("Sheet1", column, obj[field])
		}
		startRowIndex++
	}
	f.SetActiveSheet(index)

	if err := f.SaveAs(excelName); err != nil {
		return err
	}
	return nil
}


func transWord(w string,offset int32) string {
	b := []rune(w)
	for i, r := range w {
		if r >= 'A' && r <= 'Z' {
			b[i] = r + offset
		}
	}
	return string(b)
}
