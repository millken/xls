# xls

[![GoDoc](https://godoc.org/github.com/millken/xls?status.svg)](https://godoc.org/github.com/millken/xls)

Pure Golang xls library for reading Microsoft Excel 97-2003 (.xls) binary files.

**Note:** This library only supports `.xls` format (BIFF8 binary). For `.xlsx` files (Excel 2007+), use libraries like [excelize](https://github.com/qax-os/excelize).

Forked from [extrame/xls](https://github.com/extrame/xls), with improvements:
- Zero external dependencies (except `golang.org/x/text` for encoding)
- Simplified codebase (reduced from 17 to 7 source files)
- Inlined OLE2 parsing library
- Bug fixes and code optimizations

## Installation

```bash
go get github.com/millken/xls
```

## Usage

### Basic Usage

```go
package main

import (
    "fmt"
    "github.com/millken/xls"
)

func main() {
    // Open the xls file
    xlFile, err := xls.Open("test.xls", "utf-8")
    if err != nil {
        panic(err)
    }

    // Get sheet by index
    sheet := xlFile.GetSheet(0)
    if sheet == nil {
        return
    }

    // Iterate over rows
    for i := 0; i <= int(sheet.MaxRow); i++ {
        row := sheet.Row(i)
        if row == nil {
            continue
        }
        // Get cell value by column index
        fmt.Println(row.Col(0), row.Col(1))
    }
}
```

### Open with Reader

```go
file, err := os.Open("test.xls")
if err != nil {
    panic(err)
}
defer file.Close()

xlFile, err := xls.OpenReader(file, "utf-8")
if err != nil {
    panic(err)
}
// Use xlFile...
```

### Read All Cells

```go
xlFile, _ := xls.Open("test.xls", "utf-8")
// Read all cells with max row limit
rows := xlFile.ReadAllCells(10000)
for _, row := range rows {
    fmt.Println(row)
}
```

## API Reference

### Functions

- `Open(filename, charset string) (*WorkBook, error)` - Open xls file
- `OpenWithCloser(filename, charset string) (*WorkBook, io.Closer, error)` - Open with closer for manual close
- `OpenReader(reader io.ReadSeeker, charset string) (*WorkBook, error)` - Open from reader

### WorkBook Methods

- `GetSheet(num int) *WorkSheet` - Get sheet by index
- `NumSheets() int` - Get total sheet count
- `ReadAllCells(max int) [][]string` - Read all cells

### WorkSheet Methods

- `Row(i int) *Row` - Get row by index
- `Name string` - Sheet name
- `MaxRow uint16` - Maximum row number
- `Selected bool` - Is sheet selected
- `Visibility TWorkSheetVisibility` - Sheet visibility state

### Row Methods

- `Col(i int) string` - Get cell value (handles merged cells)
- `ColExact(i int) string` - Get cell value (first cell only for merged)
- `FirstCol() int` - First column index
- `LastCol() int` - Last column index

## Features

- Read Excel 97-2004 (.xls) files
- Multiple sheets support
- Merged cells support
- Date/time formatting
- Hyperlinks
- Multiple cell types (string, number, formula, blank, RK)
- Unicode support (UTF-16, Windows-1251)

## Dependencies

- `golang.org/x/text` - Character encoding support

## License

MIT License