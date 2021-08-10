package main

import (
	"docx-wordtable2excel/common"
	"docx-wordtable2excel/docx2"
	"fmt"
	"io/ioutil"
	"strings"
	"time"

	"github.com/360EntSecGroup-Skylar/excelize"
)

func main() {
	files, _ := ioutil.ReadDir("./")
	fileDataArr := make([]common.Object, 0)
	for _, f := range files {
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
	if nil != err {
		fmt.Printf("生成 excel 文件失败 \n")
	}
	fmt.Printf("生成 excel 文件 %s 成功 \n", excelName)
}

func openDocxFile(filename string) ([][]string, error) {
	defer func() {
		if err := recover(); err != nil {
			fmt.Println(err)
		}
	}()

	reader, readerFiles := docx2.UnpackDocx(filename)
	if reader == nil || len(readerFiles) == 0 {
		return nil, fmt.Errorf("empty docx reader")
	}
	defer reader.Close()
	docFile := docx2.RetrieveWordDoc(readerFiles)
	doc := docx2.OpenWordDoc(*docFile)
	content := docx2.WordDocToString(doc)
	exectContent := docx2.Extract(content)
	t := exectContent.Body.Table.TableRow

	var table [][]string
	for _, r := range t {
		var row []string
		for _, c := range r.TableColumn {
			row = append(row, strings.Join(c.Cell, ""))
		}
		table = append(table, row)
	}

	return table, nil
}

func pickColumnTableFiled(table [][]string) common.Object {
	dataMap := common.Object{}
	for _, row := range table {
		for index, field := range row {
			if newField, ok := common.FocusField2ExcelMap[field]; ok {
				if index+1 <= len(row) {
					dataMap.Set(newField, row[index+1])
				}
			}
		}
	}
	return dataMap
}

func Save2Excel(excelName string, fileDataArr []common.Object) error {
	f := excelize.NewFile()

	index := f.NewSheet(common.SheetName)

	startColumnWord := "A"
	startRowIndex := 1

	for offset, field := range common.FocusField2ExcelArr {
		tmpColumnWord := transWord(startColumnWord, int32(offset))
		column := fmt.Sprintf("%s%d", tmpColumnWord, startRowIndex)
		f.SetCellValue(common.SheetName, column, field)
	}

	startRowIndex++
	for _, obj := range fileDataArr {
		for offset, field := range common.FocusField2ExcelArr {
			tmpColumnWord := transWord(startColumnWord, int32(offset))
			column := fmt.Sprintf("%s%d", tmpColumnWord, startRowIndex)
			f.SetCellValue(common.SheetName, column, obj[field])
		}
		startRowIndex++
	}
	f.SetActiveSheet(index)

	if err := f.SaveAs(excelName); nil != err {
		return err
	}
	return nil
}

func transWord(w string, offset int32) string {
	b := []rune(w)
	for i, r := range w {
		if r >= 'A' && r <= 'Z' {
			b[i] = r + offset
		}
	}
	return string(b)
}
