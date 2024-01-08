package excel

import "reflect"

type FieldMappingItem struct {
	Name      string
	Index     int
	ExcelType reflect.Type
}
