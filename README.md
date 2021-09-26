# excel
Package excel provides simple usage for reading excel to structs 
and the sheets in the excel need to have the same structure.

```
go get github.com/tangximing/excel
```

## Scan rows to data
```
var excelPath string
f, err := NewExcelFromFile(excelPath, HeaderRow(2))
if err != nil {
    return
}

rows, err := f.GetRowsWithoutHeader()
if err != nil {
    return
}

test := new(test)
err = f.ScanRow(rows[0], test)
if err != nil {
    return
}
fmt.Println(test.Field1.GetStdValue())
```

## Async Scan rows to data
```
f, err := NewExcelFromFile("./test/test.xlsx", HeaderRow(2))
if err != nil {
    return
}

rows, err := f.GetRowsWithoutHeader()
if err != nil {
    return
}

t1 := new(test)
for row := range f.AsyncScanRows(rows, t1) {
    if row.Err != nil {
        return
    }

    t2 := row.Responses[0].(*test)
    fmt.Println(*t2)
}
```
