package excel

import "reflect"

type FieldMappingItem struct {
	// the fieldName of the struct represent the row.
	FieldName string
	// the column name of the sheet, in the other word, is the header row of the sheet.
	ColName string
	// the index of columns in the header row.
	ColIndex int
	// the fieldType
	FieldType reflect.Type
}
