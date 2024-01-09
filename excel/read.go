package excel

import (
	"errors"
	"fmt"
	"reflect"
	"strconv"
	"strings"

	log "github.com/sirupsen/logrus"
	"github.com/xuri/excelize/v2"
)

func init() {
	log.SetLevel(log.FatalLevel)
}

// Read the data from the sheet
func ReadFromSheet[T any](filepath string, sheetName string) ([]T, error) {
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
	rows, err := f.GetRows(sheetName)
	if err != nil {
		return nil, fmt.Errorf("No sheet with the specified name exists.")
	}
	if len(rows) <= 1 {
		return nil, fmt.Errorf("No data in the sheet.")
	}
	numberFormatIsUpdated := false
	dataRowCount := 0
	fieldMapping := make(map[string]*FieldMappingItem, 0)
	style, _ := f.NewStyle(&excelize.Style{NumFmt: 1})

	// read the first row, init fieldMapping
	for idx, row := range rows {
		if len(row) == 0 {
			// skip the black rows
			continue
		}
		colNameMappingIndex, err := initColNameMappingIndex(row)
		if err != nil {
			return nil, err
		}

		item := new(T)
		t := reflect.TypeOf(item)
		t = t.Elem()
		fieldNum := t.NumField()
		for i := 0; i < fieldNum; i++ {
			fieldIndexSetted := false
			fieldName := t.Field(i).Name
			fieldTags := strings.Split(t.Field(i).Tag.Get(readTag), ",")
			fieldType := t.Field(i).Type
			for key := range colNameMappingIndex {
				if containsInArray(key, fieldTags) {
					fieldMapping[fieldName] = &FieldMappingItem{
						FieldName: fieldName,
						FieldType: fieldType,
						ColIndex:  colNameMappingIndex[key],
						ColName:   key,
					}
					fieldIndexSetted = true
					delete(colNameMappingIndex, key)
					break
				}
			}
			if !fieldIndexSetted {
				return nil, fmt.Errorf("The field=%s not found in sheet header.", fieldName)
			}
			if fieldType.Kind() == reflect.Struct && fieldType.String() == "time.Time" {
				columnIndexStr, _ := excelize.ColumnNumberToName(fieldMapping[fieldName].ColIndex + 1)
				err := f.SetColStyle(sheetName, columnIndexStr, style)
				if err != nil {
					return nil, errors.New("set style for column failed. ")
				}
				numberFormatIsUpdated = true
			}
		}
		dataRowCount = len(rows) - idx - 1
		break
	}

	if numberFormatIsUpdated {
		// read the sheet again because the numberFormat is true
		rows, err = f.GetRows(sheetName)
		if err != nil {
			return nil, fmt.Errorf("can't find the sheet with the sheetName = %s\n", sheetName)
		}
	}
	results := make([]T, 0, dataRowCount)

	// reade the data
	headerHasNotBeenRead := true
	for _, row := range rows {
		if len(row) == 0 {
			// skip the black rows
			continue
		}
		if headerHasNotBeenRead {
			headerHasNotBeenRead = false
		} else {
			item := new(T)
			v := reflect.ValueOf(item)
			err = setDataForObject(v, row, fieldMapping)
			if err != nil {
				return nil, err
			}
			results = append(results, *item)
		}

	}
	return results, nil
}

// Set the object value for each row data
func setDataForObject(item reflect.Value, cells []string, fieldMapping map[string]*FieldMappingItem) error {
	if item.Type().Kind() == reflect.Pointer {
		item = item.Elem()
	}

	for k, v := range fieldMapping {
		field := item.FieldByName(k)
		kind := field.Type().Kind()
		// if kind == reflect.Pointer {
		// 	kind = field.Type().Elem().Kind()
		// }
		switch kind {
		case reflect.String:
			set2String(field, cells[v.ColIndex])
		case reflect.Int, reflect.Int32, reflect.Int64, reflect.Int8, reflect.Int16:
			err := set2Int64(field, cells[v.ColIndex])
			if err != nil {
				return fmt.Errorf("col=%s, %s", fieldMapping[k].ColName, err.Error())
			}
		case reflect.Float64, reflect.Float32:
			err := set2float64(field, cells[v.ColIndex])
			if err != nil {
				return fmt.Errorf("col=%s, %s", fieldMapping[k].ColName, err.Error())
			}
		case reflect.Bool:
			set2bool(field, cells[v.ColIndex])
		case reflect.Struct:
			if field.Type().String() == "time.Time" {
				err := set2Time(field, cells[v.ColIndex])
				if err != nil {
					return fmt.Errorf("col=%s, %s", fieldMapping[k].ColName, err.Error())
				}
			}
		case reflect.Pointer:
			return fmt.Errorf("A data field cannot defined as a pointer.")
		}

	}
	return nil
}

// Initialize the mapping between the header column name and the index of the column where the column name must be unique
func initColNameMappingIndex(cells []string) (map[string]int, error) {
	colNameMappingIndex := make(map[string]int, len(cells))
	for colIndex, cell := range cells {
		_, ok := colNameMappingIndex[cell]
		if ok {
			return nil, fmt.Errorf("The same column name exists in the sheet.")
		}
		colNameMappingIndex[cell] = colIndex
	}
	return colNameMappingIndex, nil
}

// Determines whether the tag in the field matches the header column name
func containsInArray(key string, tags []string) bool {
	key = strings.TrimSpace(key)
	for _, s := range tags {
		if strings.TrimSpace(s) == key {
			return true
		}
	}
	return false
}

// set the cell value, usually is string, to a string field
func set2String(value reflect.Value, str string) {
	switch value.Type().Kind() {
	case reflect.String:
		value.SetString(strings.TrimSpace(str))
	case reflect.Pointer:
		s := strings.TrimSpace(str)
		value.Set(reflect.ValueOf(&s))
	}
}

// set the cell value, usually is string, to a integer field
func set2Int64(value reflect.Value, str string) error {
	intValue, err := strconv.ParseInt(strings.TrimSpace(str), 0, 64)
	if err != nil {
		return fmt.Errorf("failed to convert value=%s to a int", str)
	}
	switch value.Type().Kind() {
	case reflect.Int, reflect.Int32, reflect.Int64, reflect.Int8, reflect.Int16:
		value.SetInt(intValue)
	case reflect.Pointer:
		switch value.Type().Elem().Kind() {
		case reflect.Int:
			num := int(intValue)
			value.Set(reflect.ValueOf(&num))
		case reflect.Int32:
			num := int32(intValue)
			value.Set(reflect.ValueOf(&num))
		case reflect.Int64:
			value.Set(reflect.ValueOf(&intValue))
		}
	}
	return nil
}

// set the cell value, usually is string, to a float field
func set2float64(value reflect.Value, str string) error {
	floatValue, err := strconv.ParseFloat(strings.TrimSpace(str), 64)
	if err != nil {
		return errors.New("failed to convert to a float")
	}
	switch value.Type().Kind() {
	case reflect.Float64, reflect.Float32:
		value.SetFloat(floatValue)
	case reflect.Pointer:
		value.Set(reflect.ValueOf(&floatValue))
	}
	return nil
}

// set the cell value, usually is string, to a time.Time field
func set2Time(value reflect.Value, str string) error {
	floatValue, err := strconv.ParseFloat(str, 64)
	if err != nil {
		return fmt.Errorf("failed to convert value=%s to a time", str)
	}
	toTime, err := excelize.ExcelDateToTime(floatValue, false)
	if err != nil {
		return fmt.Errorf("failed to convert value=%s to a time", str)
	}
	switch value.Type().Kind() {
	case reflect.Struct:
		value.Set(reflect.ValueOf(toTime))
	case reflect.Pointer:
		value.Set(reflect.ValueOf(&toTime))
	}
	return nil
}

// set the cell value, usually is string, to a bool field
func set2bool(value reflect.Value, str string) {
	s := strings.ToUpper(strings.TrimSpace(str))
	v := true
	if s == "FALSE" || s == "0" {
		v = false
	}
	switch value.Type().Kind() {
	case reflect.Bool:
		value.SetBool(v)
	case reflect.Pointer:
		value.Set(reflect.ValueOf(&v))
	}
}
