package main

import (
	"log"
	"fmt"

	"github.com/360EntSecGroup-Skylar/excelize"
)

func main() {
	source := read("sample.xls")

	sheets := source.GetSheetMap()
	// get title row
	rows, err := source.GetRows
}

func read(name string) *excelize.File {
	f, err := excelize.OpenFile(name)
	if err != nil {
		log.Fatalf("Error on open excel source:\n%v", err)
	}
	return f
}


