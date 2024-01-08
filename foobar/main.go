package main

import (
	"fmt"
	"reflect"
)

type User struct {
	Id   int
	Name string `json:"name1" db:"name2"`
	Age  int
}

func main() {
	var s User
	// v := reflect.ValueOf(&s)
	// t := v.Type()
	t := reflect.TypeOf(&s)
	f := t.Elem().Field(1)
	fmt.Println(f.Tag.Get("json"))
	fmt.Println(f.Tag.Get("db"))
}
