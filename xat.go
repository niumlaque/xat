package main

import (
	"errors"
	"fmt"
	"os"
	"strings"

	"github.com/jessevdk/go-flags"
	"github.com/tealeg/xlsx"
)

type options struct {
	Filename       string
	Sheetname      string
	Separator      string `short:"s" long:"separator" description:"Column separator" default:"\t"`
	PrintRowNumber bool   `long:"print-row-num" description:"Print row number"`
	PrintEmptyRow  bool   `long:"print-empty-row" description:"Print empty row"`
}

func getOptions(args []string) (*options, error) {
	opts := &options{}
	args, err := flags.ParseArgs(opts, args)
	if err != nil {
		return nil, err
	}

	if len(args) >= 1 {
		opts.Filename = args[0]
	}

	if len(args) >= 2 {
		opts.Sheetname = args[1]
	}

	return opts, nil
}

func failOnError(err error) {
	if err != nil {
		fmt.Fprintf(os.Stderr, "%s\n", err.Error())
		os.Exit(1)
	}
}

func getSheetOrHead(wb *xlsx.File, name string) (*xlsx.Sheet, error) {
	if len(wb.Sheets) == 0 {
		return nil, errors.New("No sheet in file")
	}

	var sheet *xlsx.Sheet = nil
	for _, x := range wb.Sheets {
		if x.Name == name {
			return x, nil
		}
	}

	if len(name) > 0 {
		return nil, fmt.Errorf("Sheet \"%s\" not found in file", name)
	}

	if sheet == nil {
		return wb.Sheets[0], nil
	}

	return nil, errors.New("Unexpected error occurred")
}

func outputLine(rowNum int, text string, opts *options) {
	if opts.PrintEmptyRow {
		if opts.PrintRowNumber {
			fmt.Fprintf(os.Stdout, "%d%s%s\n", rowNum, opts.Separator, text)
		} else {
			fmt.Fprintf(os.Stdout, "%s\n", text)
		}
	} else {
		trimmed := strings.TrimRight(text, opts.Separator)
		if len(trimmed) > 0 {
			if opts.PrintRowNumber {
				fmt.Fprintf(os.Stdout, "%d%s%s\n", rowNum, opts.Separator, text)
			} else {
				fmt.Fprintf(os.Stdout, "%s\n", text)
			}
		}
	}
}

func main() {
	opts, err := getOptions(os.Args[1:])
	if err != nil {
		if flags.WroteHelp(err) {
			os.Exit(1)
		} else {
			failOnError(err)
		}
	}

	if len(opts.Separator) == 0 {
		opts.Separator = "\t"
	}

	wb, err := xlsx.OpenFile(opts.Filename)
	failOnError(err)

	sheet, err := getSheetOrHead(wb, opts.Sheetname)
	failOnError(err)

	for i, row := range sheet.Rows {
		if len(row.Cells) == 0 {
			continue
		}

		texts := make([]string, 0, len(row.Cells))
		for _, cell := range row.Cells {
			texts = append(texts, cell.String())
		}

		line := strings.Join(texts, opts.Separator)
		outputLine(i, line, opts)
	}
}
