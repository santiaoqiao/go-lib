package main

import (
	"fmt"

	"github.com/santiaoqiao/go-lib/excel"
)

type JSProvience struct {
	Id         string `x-read:"序号,编号"`
	Provience  string `x-read:"省分"`
	City       string `x-read:"城市"`
	CityAlias  string `x-read:"简称"`
	Code       string `x-read:"邮政编码"`
	PhoneZCode string `x-read:"电话区号"`
	CarCode    string `x-read:"车牌号"`
	CityClass1 string `x-read:"城市分级"`
	CItyClass2 string `x-read:"城市规划分级"`
	Vender     string `x-read:"厂家"`
}

func main() {
	s, err := excel.ReadFromSheet[JSProvience]("基础信息.xlsx", "江苏城市分级", excel.Config{Trim: false})
	if err != nil {
		fmt.Println(err.Error())
	}
	for _, v := range s {
		fmt.Printf("%#v\n", v)
	}
}
