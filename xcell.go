package main

import (
	"fmt"
	"github.com/tealeg/xlsx"
	"io/ioutil"
	"os"
	"path"
	"strings"
)

const (
	maxDiffSize = 100
)

type diff struct {
	row  int
	col  int
	str1 string
	str2 string
}

func (d diff) String() string {
	return fmt.Sprintf("%d:%d | %-10s | %-10s |", d.row, d.col, d.str1, d.str2)
}

func main() {
	subCommand := os.Args[1]
	if subCommand == "diff" {
		diffFiles(os.Args[2], os.Args[3])
	} else if subCommand == "conv" {
		if len(os.Args) == 3 {
			ext := path.Ext(os.Args[2])
			dest := strings.TrimSuffix(os.Args[2], ext)
			convFile(os.Args[2], dest)
		} else {
			convFile(os.Args[2], os.Args[3])
		}
	}
}

func diffFiles(path1, path2 string) {
	file1, err := xlsx.OpenFile(path1)
	if err != nil {
		panic(err)
	}
	file2, err := xlsx.OpenFile(path2)
	if err != nil {
		panic(err)
	}

	for _, sheet1 := range file1.Sheets {
		exists := false
		fmt.Println(sheet1.Name)
		for _, sheet2 := range file2.Sheets {
			if sheet1.Name == sheet2.Name {
				diffs := compareSheets(sheet1, sheet2)
				printDiffs(diffs)
				exists = true
				break
			}
		}
		if !exists {
			fmt.Println("Only in " + path1)
		}
		fmt.Println()
	}

	for _, sheet2 := range file2.Sheets {
		exists := false
		for _, sheet1 := range file1.Sheets {
			if sheet2.Name == sheet1.Name {
				exists = true
				break
			}
		}
		if !exists {
			fmt.Println(sheet2.Name)
			fmt.Println("Only in " + path2)
			fmt.Println()
		}
	}
}

func compareSheets(s1, s2 *xlsx.Sheet) []*diff {
	result := []*diff{}
	for r, r1 := range s1.Rows {
		r2 := new(xlsx.Row)
		r2.Cells = make([]*xlsx.Cell, 0, 0)
		if r < len(s2.Rows) {
			r2 = s2.Rows[r]
		}
		result = append(result, compareRows(r, r1, r2)...)
	}
	if len(s1.Rows) < len(s2.Rows) {
		for r := len(s1.Rows); r < len(s2.Rows); r++ {
			r1 := new(xlsx.Row)
			r1.Cells = make([]*xlsx.Cell, 0, 0)
			r2 := s2.Rows[r]
			result = append(result, compareRows(r, r1, r2)...)
		}
	}
	return result
}

func compareRows(row int, r1, r2 *xlsx.Row) []*diff {
	result := []*diff{}
	for c, c1 := range r1.Cells {
		v1 := c1.Value
		v2 := ""
		if c < len(r2.Cells) {
			v2 = r2.Cells[c].Value
		}
		if v1 != v2 {
			result = append(result, &diff{row, c, v1, v2})
		}
	}
	if len(r1.Cells) < len(r2.Cells) {
		for c := len(r1.Cells); c < len(r2.Cells); c++ {
			v1 := ""
			v2 := r2.Cells[c].Value
			if v1 != v2 {
				result = append(result, &diff{row, c, v1, v2})
			}
		}
	}
	return result
}

func printDiffs(diffs []*diff) {
	if len(diffs) > maxDiffSize {
		fmt.Println("Too many diffs!")
	} else {
		for _, d := range diffs {
			fmt.Println(d)
		}
	}
}

func convFile(path1, path2 string) {
	info, err := os.Stat(path2)
	if err != nil || !info.IsDir() {
		os.Mkdir(path2, 600)
	}
	name2csv := toCsvs(path1)
	for name, csv := range name2csv {
		dest := path.Join(path2, name)
		ioutil.WriteFile(dest, csv, 600)
	}
}

func toCsvs(path string) map[string][]byte {
	file, err := xlsx.OpenFile(path)
	if err != nil {
		panic(err)
	}

	result := make(map[string][]byte)
	for _, sheet := range file.Sheets {
		result[sheet.Name+".csv"] = toCsv(sheet)
	}
	return result
}

func toCsv(sheet *xlsx.Sheet) []byte {
	str := ""
	for _, row := range sheet.Rows {
		vals := make([]string, len(row.Cells), len(row.Cells))
		for i, cell := range row.Cells {
			vals[i] = cell.Value
		}
		str = str + strings.Join(vals, ",") + "\n"
	}
	return []byte(str)
}
