package excel

import (
	"fmt"
	"reflect"

	log "github.com/sirupsen/logrus"
	"github.com/xuri/excelize/v2"
)

func init() {
	log.SetLevel(log.TraceLevel)
}

func ReadFromSheet[T any](filepath string, sheetName string, config Config) ([]T, error) {
	var StringTrim = true
	f, err := excelize.OpenFile(filepath)
	if err != nil {
		return nil, fmt.Errorf("file opening failed. %s\n", filepath)
	}
	defer func() {
		log.Tracef("the defer function fired, the xlsx file will be closed")
		if err = f.Close(); err != nil {
			log.Fatalf("there is a mistake when file close.")
		}
	}()
	StringTrim = config.Trim
	log.Trace(StringTrim)

	//读取第一行  标题行
	rows, err := f.GetRows(sheetName)
	if err != nil {
		return nil, fmt.Errorf("No sheet with the specified name exists.")
	}
	if len(rows) <= 1 {
		return nil, fmt.Errorf("No data in the sheet.")
	}
	// 创建容器
	results := make([]T, len(rows))
	// headMapping := make(map[string]int, 0)
	var headerMapping map[string]HeaderMappingItem

	for idx := range rows {
		if idx == 0 {
			// 第一行，建立字段对应的索引
			headerMapping = make(map[string]HeaderMappingItem, 20)

		} else {
			item := new(T)
			v := reflect.ValueOf(item)
			// t := reflect.TypeOf(item)
			// fmt.Printf("v: %v\n", v.Elem().NumField())
			for i := 0; i < v.Elem().NumField(); i++ {
				ft := v.Elem().Type().Field(i)
				headerMapping[ft.Name] = HeaderMappingItem{
					name:      ft.Name,
					excelType: ft.Type,
				}
			}

			results = append(results, *item)
		}
	}
	// 建立字段的对应的索引

	// 建立 slice

	// 循环读取

	return results, nil
}

func checkSheetName(f *excelize.File, sheetName *string) (int, error) {
	s := f.GetSheetList()
	for index, name := range s {
		if name == *sheetName {
			return index, nil
		}
	}
	return -1, fmt.Errorf("No sheet with the specified name exists.")
}
