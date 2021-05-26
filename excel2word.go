package main

import (
	"docx-wordtable2excel/common"
	"fmt"
	"io/ioutil"
	"strings"

	"github.com/nguyenthenguyen/docx"

	"github.com/360EntSecGroup-Skylar/excelize"
)

func main() {
	files, _ := ioutil.ReadDir("./")
	excelFileName := ""
	demoWordFileName := ""
	for _, f := range files {
		if f.IsDir() {
			continue
		}

		fileSplits := strings.Split(f.Name(), ".")
		if len(fileSplits) != 2 {
			continue
		}

		if fileSplits[1] == "xlsx" {
			excelFileName = f.Name()
		}
		if f.Name() == common.DemoWordFileName {
			demoWordFileName = f.Name()
		}
	}
	if excelFileName == "" || demoWordFileName == "" {
		fmt.Println("没有找到对应 excel 或者 demo.docx 文件")
		return
	}

	excelRowDatas, err := pickExcelColumn(excelFileName)
	if nil != err {
		fmt.Printf("获取 excel 行元素失败，错误：%v \n", err)
		return
	}

	demoF, err := docx.ReadDocxFile(demoWordFileName)
	if nil != err {
		fmt.Printf("生成 word 文件失败，错误：%v\n", err)
		return
	}
	defer demoF.Close()

	for _, row := range excelRowDatas {
		tmpDocx := demoF.Editable()
		docxFileName := "回复工单--"
		for _, column := range row {
			if column.Ignore {
				docxFileName = fmt.Sprintf("%s_%s", docxFileName, column.ColumnVal)
				continue
			}

			_ = tmpDocx.Replace(column.ReplaceField, column.ColumnVal, -1)
		}
		err = tmpDocx.WriteToFile("./tmp/" + docxFileName)
	}

}

func pickExcelColumn(fileName string) ([][]*common.FieldItem, error) {
	f, err := excelize.OpenFile(fileName)
	if err != nil {
		return nil, err
	}

	sheetMap := f.GetSheetMap()
	excelRowDatas := make([][]*common.FieldItem, 0)

	for _, sheet := range sheetMap {
		rows := f.GetRows(sheet)
		excelIndexMap := make(map[string]*common.FieldItem)
		for rIndex, row := range rows {
			if rIndex == 0 {
				if excelIndexMapTmp, has := matchFocusColumnIndex(row); has {
					excelIndexMap = excelIndexMapTmp
					continue
				} else {
					return nil, fmt.Errorf("excel: %s 中没有匹配到所需的字段")
				}
			}

			tmpRowDatas := make([]*common.FieldItem, 0)
			for _, fItem := range excelIndexMap {
				tmpRowItem := &common.FieldItem{
					Ignore:       fItem.Ignore,
					KeyField:     fItem.KeyField,
					ReplaceField: fItem.ReplaceField,
					ColumnIndex:  fItem.ColumnIndex,
				}
				if fItem.Ignore {
					tmpRowItem.ColumnVal = fmt.Sprintf("%s_%s", sheet, tmpRowItem.ColumnVal)
				}

				// 兼容老文件
				columnVal := ""
				if fItem.ColumnIndex >= 0 {
					columnVal = row[fItem.ColumnIndex]
				}
				tmpRowItem.ColumnVal = columnVal

				tmpRowDatas = append(tmpRowDatas, tmpRowItem)
			}
			excelRowDatas = append(excelRowDatas, tmpRowDatas)
		}
	}
	return excelRowDatas, nil
}

func matchFocusColumnIndex(row []string) (map[string]*common.FieldItem, bool) {
	excel2WordFileMap := common.NewExcel2WordFileMap()
	has := false
	for index, v := range row {
		if fItem, ok := excel2WordFileMap[v]; ok {
			has = true
			fItem.ColumnIndex = index
		}
	}
	return excel2WordFileMap, has
}
