package main

import (
	"flag"
	"fmt"
	"os"

	"github.com/tealeg/xlsx"
)

func main() {
	var (
		in  = flag.String("in", "", "input xlsx file")
		out = flag.String("out", "", "output xlsx file")
	)
	flag.Parse()
	if len(*in) == 0 || len(*out) == 0 {
		fmt.Fprintf(os.Stderr, "missing in or out flags\n")
		flag.PrintDefaults()
		os.Exit(1)
	}
	xlFile, err := xlsx.OpenFile(*in)
	if err != nil {
		fmt.Fprintf(os.Stderr, "Failed to open: %v\n", err)
		os.Exit(2)
	}
	for _, sheet := range xlFile.Sheets {
		for _, row := range sheet.Rows {
			for _, cell := range row.Cells {
				if len(cell.Formula()) > 0 {
					cell.Value = cell.String()
					cell.SetFormula("")
				}

			}
		}
	}
	err = xlFile.Save(*out)
	if err != nil {
		fmt.Fprintf(os.Stderr, "Failed to save: %v\n", err)
		os.Exit(2)
	}
}
