package excel

import (
	"fmt"
	"os"
	"testing"
	"time"

	"github.com/stretchr/testify/assert"
)

type test struct {
	Field1 StringField `excel:"字段|字段1"`
	Field2 IntField    `excel:"字段|字段2"`
	Field3 BoolField   `excel:"字段|字段3"`
	Field4 TimeField   `excel:"字段|字段4"`
	Field5 FloatField  `excel:"字段|字段5"`
}

func TestNewExcelFromData(t *testing.T) {
	var tests []interface{}
	for i := 1; i <= 5; i++ {
		tests = append(tests, &test{
			Field1: NewStringField(fmt.Sprintf("%d", i)),
			Field2: NewIntField(i),
			Field3: NewBoolField(true),
			Field4: NewTimeField(time.Now().UTC()),
			Field5: NewFloatField(1.2),
		})
	}

	f, err := NewExcelFromData(tests, SheetCount(2), SheetPrefix("Sheet"), HeaderRow(2))
	assert.Nil(t, err)

	exelFileName := "test.xlsx"
	err = f.GetFile().SaveAs(exelFileName)
	assert.Nil(t, err)
	_ = os.Remove(exelFileName)
}

func TestExcel_IsHeaderConsistent(t *testing.T) {
	f, err := NewExcelFromFile("./test/test.xlsx", HeaderRow(2))
	assert.Nil(t, err)

	test := new(test)

	isConsistent, err := f.IsHeaderConsistent(test)
	assert.Nil(t, err)
	assert.True(t, isConsistent)
}

func TestExcel_ScanRow(t *testing.T) {
	f, err := NewExcelFromFile("./test/test.xlsx", HeaderRow(2))
	assert.Nil(t, err)

	rows, err := f.GetRowsWithoutHeader()
	assert.Nil(t, err)

	test := new(test)

	err = f.ScanRow(rows[0], test)
	assert.Nil(t, err)
	fmt.Println(test)
	assert.Equal(t, test.Field1.GetStdValue(), "1")
}

func TestExcel_AsyncScanRows(t *testing.T) {
	f, err := NewExcelFromFile("./test/test.xlsx", HeaderRow(2))
	assert.Nil(t, err)

	rows, err := f.GetRowsWithoutHeader()
	assert.Nil(t, err)

	t1 := new(test)
	for row := range f.AsyncScanRows(rows, t1) {
		assert.Nil(t, row.Err)
		t2 := row.Responses[0].(*test)
		fmt.Println(*t2)
	}
}
