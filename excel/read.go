package excel

import (
	"fmt"
	"reflect"
	"strconv"
	"strings"

	log "github.com/sirupsen/logrus"
	"github.com/xuri/excelize/v2"
)

func init() {
	log.SetLevel(log.TraceLevel)
}

func ReadFromSheet[T any](filepath string, sheetName string, config Config) ([]T, error) {
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
	// StringTrim = config.Trim
	//读取第一行  标题行
	rows, err := f.GetRows(sheetName)
	if err != nil {
		return nil, fmt.Errorf("No sheet with the specified name exists.")
	}
	if len(rows) <= 1 {
		return nil, fmt.Errorf("No data in the sheet.")
	}
	// 创建容器
	results := make([]T, 0, len(rows)-1)
	// log.Trace(len(rows))
	var fieldMapping map[string]*FieldMappingItem
	for idx, cells := range rows {
		if idx == 0 {
			// 第一行，建立字段对应的索引
			item := new(T)
			fieldMapping = make(map[string]*FieldMappingItem, 0)
			v := reflect.ValueOf(item)
			err := initHeaderMapping(fieldMapping, v, cells)
			if err != nil {
				return nil, err
			}
			// for key, val := range fieldMapping {
			// 	fmt.Printf("%s: %#v\n", key, val)
			// }

		} else {
			item := new(T)
			v := reflect.ValueOf(item)
			err := setDataForObject(v, cells, fieldMapping)
			if err != nil {
				return nil, fmt.Errorf("at row=%d, %s", idx+1, err.Error())
			}
			results = append(results, *item)
		}
	}

	// for _, v := range results {
	// 	fmt.Printf("%v\n", v)
	// }

	return results, nil
}

func setDataForObject(v reflect.Value, cells []string, fieldMapping map[string]*FieldMappingItem) error {
	if v.Type().Kind() == reflect.Pointer {
		v = v.Elem()
	}
	for i := 0; i < v.NumField(); i++ {
		key := v.Type().Field(i).Name
		kind := fieldMapping[key].ExcelType.Kind()
		cellVal := cells[fieldMapping[key].Index]
		switch kind {
		case reflect.Float32, reflect.Float64:
			f, err := strconv.ParseFloat(cellVal, 64)
			if err != nil {
				return fmt.Errorf("failed to convert value:%s to a float.", cellVal)
			}
			v.Field(i).SetFloat(f)
		case reflect.Int64, reflect.Int, reflect.Int16, reflect.Int32:
			iv, err := strconv.ParseInt(cellVal, 10, 64)
			if err != nil {
				return fmt.Errorf("failed to convert value:%s to a float.", cellVal)
			}
			v.Field(i).SetInt(iv)
		case reflect.String:
			v.Field(i).SetString(cellVal)
		}
	}
	return nil
}

func initHeaderMapping(fieldMapping map[string]*FieldMappingItem, v reflect.Value, cells []string) error {
	colMappingIndex := make(map[string]int, len(cells))
	for colIndex, cell := range cells {
		_, ok := colMappingIndex[cell]
		if ok {
			return fmt.Errorf("The same column name exists in the sheet.")
		}
		colMappingIndex[cell] = colIndex
	}

	for i := 0; i < v.Elem().NumField(); i++ {
		fieldIndexSetted := false
		ft := v.Elem().Type().Field(i)
		tags := strings.Split(ft.Tag.Get(readTag), ",")
		for key := range colMappingIndex {
			if containsInArray(key, tags) {
				fieldMapping[ft.Name] = &FieldMappingItem{
					Name:      ft.Name,
					ExcelType: ft.Type,
					Index:     colMappingIndex[key],
				}
				fieldIndexSetted = true
				delete(colMappingIndex, key)
				break
			}
		}
		if !fieldIndexSetted {
			return fmt.Errorf("The field=%s not found in sheet header.", ft.Name)
		}
	}
	return nil
}

func containsInArray(key string, tags []string) bool {
	key = strings.TrimSpace(key)
	for _, s := range tags {
		if strings.TrimSpace(s) == key {
			return true
		}
	}
	return false
}
