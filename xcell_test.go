package main

import (
	"github.com/tealeg/xlsx"
	"os"
	"path/filepath"
	"strings"
	"testing"
)

func beforeConv() {
	file := xlsx.NewFile()
	_, err := file.AddSheet("Sheet1")
	if err != nil {
		panic(err)
	}
	_, err = file.AddSheet("Sheet2")
	if err != nil {
		panic(err)
	}
	file.Save("fixture.xlsx")
}

func afterConv() {
	os.Remove("fixture.xlsx")
	os.RemoveAll("fixture")
}

func TestConv(t *testing.T) {
	beforeConv()
	defer afterConv()
	convFile("fixture.xlsx", "fixture")
	names, err := filepath.Glob(filepath.Join("fixture", "*"))
	if err != nil {
		panic(err)
	}
	if !strings.HasSuffix(names[0], "Sheet1.csv") {
		t.Errorf("%s not found", "Sheet1.csv")
	}
	if !strings.HasSuffix(names[1], "Sheet2.csv") {
		t.Errorf("%s not found", "Sheet2.csv")
	}
}
