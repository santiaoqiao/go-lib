package tmp

import (
	"errors"
	"fmt"
	"reflect"
	"strconv"
	"strings"

	log "github.com/sirupsen/logrus"
	"github.com/xuri/excelize/v2"
)

var StringTrim = true

// rowMateInfo used in fieldMapping
type rowMateInfo struct {
	// A field in a customer struct that need to carry some mate info.
	// It will be used in map[string]rowMateInfo, the key represents field name.
	columnIndex int      // index of column in the sheet.
	alias       []string // field corresponds to a possible set of column names.
}

// Read is the main function to read Excel file
func Read[T any](fileName string, trim ...bool) (*T, map[string][]string, error) {
	//#region open the Excel file``
	f, err := excelize.OpenFile(fileName)
	if err != nil {
		return nil, nil, fmt.Errorf("file opening failed. %s\n", fileName)
	}
	defer func() {
		log.Tracef("the defer function fired, the xlsx file will be closed")
		if err = f.Close(); err != nil {
			log.Fatalf("there is a mistake when file close.")
		}
	}()
	//#endregion
	//#region indicate whether the string in the cell should use strings.trim() to erase the blank
	if len(trim) > 0 {
		StringTrim = trim[0]
	}
	//#endregion
	//#region read data
	result := new(T)
	errLogs := make(map[string][]string)

	// build a mapping structure
	v := reflect.ValueOf(result).Elem()
	t := reflect.TypeOf(result).Elem()
	fieldNum := t.NumField()
	//style, _ := f.NewStyle(`{"number_format": 0}`)
	style, _ := f.NewStyle(&excelize.Style{NumFmt: 1})
	for i := 0; i < fieldNum; i++ {
		field := t.Field(i) // is slice as []excel.Sheet1, []excel.Sheet2 ...
		if field.Type.Kind() != reflect.Slice {
			return nil, nil, fmt.Errorf("the type should be a slice, the current type is %s", field.Type.String())
		}
		currentType := field.Type.Elem() // is type as excel.Sheet1 ...
		dataRowCount := 0
		//#region read the first row as the header and init rowMateInfo
		tag := field.Tag.Get(sheetTag)
		numberFormatIsUpdated := false
		sheetName, err := getRightSheetName(f, tag)
		if err != nil {
			return nil, nil, errors.New(err.Error())
		}
		fieldMapping := initFieldMapping(field.Type.Elem())
		// get sheet data for get the first row as the header
		sheetRows, err := f.GetRows(sheetName)
		if err != nil {
			return nil, nil, fmt.Errorf("can't find the sheet with the sheetName = %s\n", sheetName)
		}
		// get the header and set the rowMateInfo
		for idx, row := range sheetRows {
			if len(row) == 0 {
				// skip the black rows
				continue
			}
			// the first row considered to be the header. read it and set the columnIndex value of the fieldMapping
			setFieldMappingWithHeader(row, fieldMapping)
			numField := currentType.NumField()
			for i := 0; i < numField; i++ {
				fieldInSheetMateInfo := currentType.Field(i)
				if (fieldInSheetMateInfo.Type.Kind() == reflect.Struct && fieldInSheetMateInfo.Type.String() == "time.Time") ||
					(fieldInSheetMateInfo.Type.Kind() == reflect.Pointer && fieldInSheetMateInfo.Type.Elem().String() == "time.Time") {
					columnIndexStr, _ := excelize.ColumnNumberToName(fieldMapping[fieldInSheetMateInfo.Name].columnIndex + 1)
					err := f.SetColStyle(sheetName, columnIndexStr, style)
					if err != nil {
						return nil, nil, errors.New("set style for column failed. ")
					}
					numberFormatIsUpdated = true
				}
			}
			dataRowCount = len(sheetRows) - idx - 1
			break
		}

		// print the fieldMapping (as rowMateInfo) for debug
		log.Debugf("sheet = %v ", sheetName)
		for k, v := range fieldMapping {
			log.Debugf("\t%v => %v\n", k, v)
		}
		log.Debugf("\n")
		//#endregion

		//#region read data for each sheet
		if numberFormatIsUpdated {
			// read the sheet again because the numberFormat is true
			sheetRows, err = f.GetRows(sheetName)
			if err != nil {
				return nil, nil, fmt.Errorf("can't find the sheet with the sheetName = %s\n", sheetName)
			}
		}
		// make a slice for the current sheet, the length & capacity is satisfied to data row count
		sheetData := reflect.MakeSlice(field.Type, dataRowCount, dataRowCount)
		errLogs[sheetName] = readData(sheetRows, fieldMapping, field.Type, sheetData)
		v.Field(i).Set(sheetData)
		//#endregion
	}
	//#endregion
	return result, errLogs, nil
}

// skip the black and header row, read the data row. used in Read() function.
func readData(rows [][]string, fieldMapping map[string]*rowMateInfo, t reflect.Type, sheetData reflect.Value) []string {
	errMsg := make([]string, 0)
	headerHasNotBeenRead := true
	dataRowIdx := 0
	for idx, row := range rows {
		// skip the black rows
		if len(row) == 0 {
			continue
		}
		if headerHasNotBeenRead {
			// the first row considered to be the header. read it and set the columnIndex value of the fieldMapping
			headerHasNotBeenRead = false
		} else {
			readData2Slice(idx, dataRowIdx, row, fieldMapping, t.Elem(), sheetData, &errMsg)
			dataRowIdx++
		}
	}
	return errMsg
}

// read a row of a sheet ro a slice. used in readData() function.
func readData2Slice(rowIdx int, dataRowIdx int, dataRow []string, fieldMapping map[string]*rowMateInfo, t reflect.Type, result reflect.Value, errMsg *[]string) {
	item := reflect.New(t)
	for k, v := range fieldMapping {
		field := item.Elem().FieldByName(k)
		cellName, _ := excelize.CoordinatesToCellName(v.columnIndex+1, rowIdx+1)
		if field.CanSet() && len(dataRow) > v.columnIndex {
			switch field.Type().Kind() {
			case reflect.String:
				set2String(field, dataRow[v.columnIndex])
			case reflect.Int, reflect.Int32, reflect.Int64, reflect.Int8, reflect.Int16:
				err := set2Int64(field, dataRow[v.columnIndex])
				if err != nil {
					*errMsg = append(*errMsg, fmt.Sprintf(err.Error()+" @ %s", cellName))
				}
			case reflect.Float64, reflect.Float32:
				err := set2float64(field, dataRow[v.columnIndex])
				if err != nil {
					*errMsg = append(*errMsg, fmt.Sprintf(err.Error()+" @ %s", cellName))
				}
			case reflect.Bool:
				set2bool(field, dataRow[v.columnIndex])
			case reflect.Struct:
				if field.Type().String() == "time.Time" {
					err := set2Time(field, dataRow[v.columnIndex])
					if err != nil {
						*errMsg = append(*errMsg, fmt.Sprintf(err.Error()+" @ %s", cellName))
					}
				}
			case reflect.Pointer: // if the field type is pointer
				switch field.Type().Elem().Kind() {
				case reflect.String:
					set2String(field, dataRow[v.columnIndex])
				case reflect.Int, reflect.Int32, reflect.Int64, reflect.Int8, reflect.Int16:
					err := set2Int64(field, dataRow[v.columnIndex])
					if err != nil {
						*errMsg = append(*errMsg, fmt.Sprintf(err.Error()+" @ %s", cellName))
					}
				case reflect.Float64, reflect.Float32:
					err := set2float64(field, dataRow[v.columnIndex])
					if err != nil {
						*errMsg = append(*errMsg, fmt.Sprintf(err.Error()+" @ %s", cellName))
					}
				case reflect.Bool:
					set2bool(field, dataRow[v.columnIndex])
				case reflect.Struct:
					if field.Type().Elem().String() == "time.Time" {
						err := set2Time(field, dataRow[v.columnIndex])
						if err != nil {
							*errMsg = append(*errMsg, fmt.Sprintf(err.Error()+" @ %s", cellName))
						}
					}
				}
			}
		}
	}
	result.Index(dataRowIdx).Set(item.Elem())
}

// use to convert the tag in the model tage to a right sheet name
func getRightSheetName(f *excelize.File, sheetName string) (string, error) {
	sheetCount := f.SheetCount
	if sheetName[0] == '[' && sheetName[len(sheetName)-1] == ']' {
		indexStr := strings.TrimSpace(sheetName[1 : len(sheetName)-1])
		if indexStr == "" {
			// if sheet tag declared as '[]', set the first sheet as the default sheetName
			sheetName = f.GetSheetName(0)
		} else {
			index, err := strconv.ParseInt(indexStr, 0, 0)
			if err != nil {
				return "", errors.New("the sheet tag declared in '[]' is not a number. ")
			}
			sheetIndex := int(index)
			if sheetIndex >= sheetCount {
				return "", errors.New("the sheet tag declared in '[n]' is out of the sheet count. ")
			}
			sheetName = f.GetSheetName(int(index))
		}
	}
	return sheetName, nil
}

// build a mapping between fields of customer struct and column name of a sheet
func initFieldMapping(t reflect.Type) map[string]*rowMateInfo {
	//	Args
	//		t: the customer struct type
	//	Returns
	//		map[string]rowMateInfo:
	//		the columnIndex is set to -1, indicates that it has not been set.
	//		e.g.
	//			{{"id": {columnIndex: -1,
	//					dataKind: reflect.string,
	//					alias:["id","ID","编号"]
	//					}}, ...}
	fieldMapping := make(map[string]*rowMateInfo)
	for i := 0; i < t.NumField(); i++ {
		f := t.Field(i)
		tag := f.Tag.Get(readTag)
		if tag == "" {
			continue
		}
		fmi := new(rowMateInfo)
		fmi.columnIndex = -1
		fmi.alias = strings.Split(tag, ",")
		fieldMapping[f.Name] = fmi
	}
	return fieldMapping
}

// update columnIndex field of the fieldMapping's every value.
func setFieldMappingWithHeader(firstRow []string, fieldMapping map[string]*rowMateInfo) {
	for _, v := range fieldMapping {
		for i, str := range firstRow {
			if isInTheSlice(str, v.alias) && v.columnIndex == -1 {
				v.columnIndex = i
				continue
			}
		}
	}
}

// Determines if the string in the slice
func isInTheSlice(str string, slice []string) bool {
	for _, name := range slice {
		if name == str {
			return true
		}
	}
	return false
}

// set the cell value, usually is string, to a string field
func set2String(value reflect.Value, str string) {
	switch value.Type().Kind() {
	case reflect.String:
		if StringTrim {
			value.SetString(strings.TrimSpace(str))
		} else {
			value.SetString(str)
		}
	case reflect.Pointer:
		if StringTrim {
			s := strings.TrimSpace(str)
			value.Set(reflect.ValueOf(&s))
		} else {
			value.Set(reflect.ValueOf(&str))
		}
	}
}

// set the cell value, usually is string, to a integer field
func set2Int64(value reflect.Value, str string) error {
	intValue, err := strconv.ParseInt(strings.TrimSpace(str), 0, 64)
	if err != nil {
		return errors.New("failed to convert to a int")
	}
	switch value.Type().Kind() {
	case reflect.Int, reflect.Int32, reflect.Int64, reflect.Int8, reflect.Int16:
		value.SetInt(intValue)
	case reflect.Pointer:
		value.Set(reflect.ValueOf(&intValue))
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
	floatValue, err := strconv.ParseFloat(strings.TrimSpace(str), 64)
	if err != nil {
		return errors.New("failed to convert to a time")
	}
	toTime, err := excelize.ExcelDateToTime(floatValue, false)
	if err != nil {
		return errors.New("failed to convert to a time")
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
