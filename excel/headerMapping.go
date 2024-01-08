package excel

import "reflect"

type HeaderMappingItem struct {
	name      string
	index     int
	excelType reflect.Type
}
