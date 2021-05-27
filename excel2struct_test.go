package main

import (
	"fmt"
	"io"
	"io/ioutil"
	"net/http"
	"testing"
)

type S struct {
	ID    int      `json:"id"`
	Level int      `json:"level"`
	Name  string   `json:"name"`
	Min   int      `json:"min"`
	Max   int      `json:"max"`
	Arr   []string `json:"arr"`
}

func TestFile2Struct(t *testing.T) {
	addr := "./file/test_config.xlsx"
	sArr := make([]S, 0)
	err2 := ParseExcelByDir(addr, "Sheet1", &sArr)
	fmt.Println(err2)
	for i := range sArr {
		fmt.Printf("res:%+v\n", sArr[i])
	}
}
func TestHttp2Struct(t *testing.T) {
	addr := "http://img-ys011.didistatic.com/static/xxl/do1_PEJ2GaPpTsbqrd7TgrlA"
	resp, err := http.Get(addr)
	if err != nil {
		fmt.Println(err)
	}
	defer func(body io.ReadCloser) {
		_ = body.Close()
	}(resp.Body)
	sArr := make([]S, 0)
	data, _ := ioutil.ReadAll(resp.Body)
	err2 := ParseExcelByBytes(data, "Sheet1", &sArr)
	fmt.Println(err2)
	for i := range sArr {
		fmt.Printf("res:%+v\n", sArr[i])
	}
}
