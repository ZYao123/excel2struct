package main

import (
	"errors"
	"fmt"
	"reflect"
	"strconv"
	"strings"
	"time"

	"github.com/tealeg/xlsx"
)

const skipRow = 3
const tagName = "json"
const splitMark = "|"

var (
	timeType = reflect.TypeOf(time.Time{})
)

func ParseExcelByDir(dir string, sheetName string, ptr interface{}) error {
	xlFile, err := xlsx.OpenFile(dir)
	if err != nil {
		return err
	}
	return parseXlsx(xlFile, sheetName, ptr)
}

func ParseExcelByBytes(file []byte, sheetName string, ptr interface{}) (err error) {
	defer func() {
		if err2 := recover(); err2 != nil {
			err = fmt.Errorf("parse xlsx fail %v", err2)
		}
	}()

	xlFile, err := xlsx.OpenBinary(file)
	if err != nil {
		return err
	}
	return parseXlsx(xlFile, sheetName, ptr)
}
func parseXlsx(xlFile *xlsx.File, sheetName string, ptr interface{}) error {
	rtPtr := reflect.TypeOf(ptr)
	if rtPtr.Kind() != reflect.Ptr {
		return errors.New("type error, is not ptr")
	}
	slice := rtPtr.Elem()
	if slice.Kind() != reflect.Slice {
		return errors.New("type error, is not slice")
	}
	rt := slice.Elem() // slice type
	newArr := make([]reflect.Value, 0)
	// 遍历sheet页读取
	for _, sheet := range xlFile.Sheets {
		if sheetName != "" && sheet.Name != sheetName {
			continue
		}
		// excel中key的cell列数
		indexMap := make(map[string]int)
		//	第一行中文描述，第二行key索引
		row1 := sheet.Rows[1]
		for i := range row1.Cells {
			indexMap[row1.Cells[i].String()] = i
		}
		// 第三行开始数据 遍历行读取
		for i := skipRow; i < len(sheet.Rows); i++ {
			row := sheet.Rows[i]
			if len(row.Cells) == 0 {
				continue
			}
			node := reflect.New(rt).Elem()
			skip := true
			for j := 0; j < rt.NumField(); j++ {
				//	要求tag
				if key, hasTag := rt.Field(j).Tag.Lookup(tagName); hasTag {
					if cellIndex, hasKey := indexMap[key]; hasKey {
						text := row.Cells[cellIndex].String()
						if !node.Field(j).CanSet() || text == "" {
							continue
						}
						skip = false
						switch node.Field(j).Kind() {
						case reflect.Bool:
							parseBool, _ := strconv.ParseBool(text)
							node.Field(j).SetBool(parseBool)
						case reflect.Int, reflect.Int8, reflect.Int16, reflect.Int32, reflect.Int64:
							atoi, _ := strconv.ParseInt(text, 10, 64)
							node.Field(j).SetInt(atoi)
						case reflect.Uint, reflect.Uintptr, reflect.Uint8, reflect.Uint16, reflect.Uint32, reflect.Uint64:
							atoi, _ := strconv.ParseUint(text, 10, 64)
							node.Field(j).SetUint(atoi)
						case reflect.String:
							node.Field(j).SetString(text)
						case reflect.Slice:
							if cArr, err := setArr(text, node, j); err == nil {
								node.Field(j).Set(reflect.Append(node.Field(j), cArr...))
							} else {
								return err
							}
						case reflect.Struct:
							if node.Field(j).Type() == timeType {
								if t, err := parseTime(text); err == nil {
									node.Field(j).Set(reflect.ValueOf(t))
								} else {
									return err
								}
							}
						default:
							return errors.New("unsupported type:" + rt.Field(j).Name + " " + node.Field(j).Kind().String())
						}
					}
				}
			}
			if !skip {
				newArr = append(newArr, node)
			}
		}
		rv := reflect.ValueOf(ptr).Elem()
		rv.Set(reflect.Append(rv, newArr...))
		return nil
	}
	return errors.New("sheet not exist")
}

func setArr(text string, node reflect.Value, j int) ([]reflect.Value, error) {
	split := strings.Split(text, splitMark)
	cArr := make([]reflect.Value, 0)
	for i2 := range split {
		nd := reflect.New(node.Field(j).Type().Elem()).Elem()
		switch nd.Kind() {
		case reflect.Bool:
			parseBool, _ := strconv.ParseBool(text)
			nd.SetBool(parseBool)
		case reflect.Int, reflect.Int8, reflect.Int16, reflect.Int32, reflect.Int64:
			parseInt, err := strconv.ParseInt(split[i2], 10, 64)
			if err != nil {
				return nil, err
			}
			nd.SetInt(parseInt)
		case reflect.Uint, reflect.Uintptr, reflect.Uint8, reflect.Uint16, reflect.Uint32, reflect.Uint64:
			atoi, _ := strconv.ParseUint(text, 10, 64)
			nd.SetUint(atoi)
		case reflect.String:
			nd.SetString(split[i2])
		}
		cArr = append(cArr, nd)
	}
	return cArr, nil
}

func parseTime(timeStr string) (time.Time, error) {
	l := len(timeStr)
	if l == len("2006-01-02") { //	"2006-01-02"	(00:00:00)	默认0点
		return time.ParseInLocation("2006-01-02", timeStr, time.Local)
	} else if l == len("2006-01-02 15:04:05") { //	"2006-01-02 15:04:05"
		return time.ParseInLocation("2006-01-02 15:04:05", timeStr, time.Local)
	} else {
		return time.Time{}, fmt.Errorf("parse time err:%s", timeStr)
	}
}
