package main

import (
	`errors`
	`github.com/tealeg/xlsx`
	`reflect`
	`strconv`
)

func ParseExcelByDir(dir string, sheetName string, ptr interface{}) error {
	xlFile, err := xlsx.OpenFile(dir)
	if err != nil {
		return err
	}
	return parseExcel(xlFile, sheetName, ptr)
}
func ParseExcelByBytes(file []byte, sheetName string, ptr interface{}) error {
	xlFile, err := xlsx.OpenBinary(file)
	if err != nil {
		return err
	}
	return parseExcel(xlFile, sheetName, ptr)
}
func parseExcel(xlFile *xlsx.File, sheetName string, ptr interface{}) error {
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
		for i := 2; i < len(sheet.Rows); i++ {
			row := sheet.Rows[i]
			node := reflect.New(rt).Elem()
			for j := 0; j < rt.NumField(); j++ {
				//	要求json tag
				if key, hasTag := rt.Field(j).Tag.Lookup("json"); hasTag {
					if cellIndex, hasKey := indexMap[key]; hasKey {
						text := row.Cells[cellIndex].String()
						if !node.Field(j).CanSet() {
							continue
						}
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
						default:
							return errors.New("unsupported type:" + rt.Field(j).Name + " " + node.Field(j).Kind().String())
						}
					}
				}
			}
			newArr = append(newArr, node)
		}
		rv := reflect.ValueOf(ptr).Elem()
		rv.Set(reflect.Append(rv, newArr...))
		return nil
	}
	return errors.New("sheet not exist")
}
