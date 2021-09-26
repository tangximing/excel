package excel

import (
	"strconv"
	"time"
)

const (
	_dateLayout = "2006-01-02"
)

type Field interface {
	// Translate trans a excel cell value (string) to a specific type value
	Translate(value string, colIndex int) (interface{}, error)
	// ColIndex get a filed col index in excel
	ColIndex() int
	// GetValue return the original value of the field
	GetValue() interface{}
}

type IntField struct {
	value    int
	colIndex int
}

func NewIntField(value int) IntField {
	return IntField{value: value}
}

var _ Field = (*IntField)(nil)

func (iField IntField) Translate(value string, colIndex int) (interface{}, error) {
	if value == "" {
		return IntField{value: 0, colIndex: colIndex}, nil
	}

	res, err := strconv.Atoi(value)
	if err != nil {
		return nil, err
	}
	return IntField{value: res, colIndex: colIndex}, nil
}

func (iField IntField) ColIndex() int {
	return iField.colIndex
}

func (iField IntField) GetStdValue() int64 {
	return int64(iField.value)
}

func (iField *IntField) SetValue(value int) {
	iField.value = value
}

func (iField IntField) GetValue() interface{} {
	return iField.value
}

type Int64Field struct {
	value    int64
	colIndex int
}

func NewInt64Field(value int64) Int64Field {
	return Int64Field{value: value}
}

var _ Field = (*Int64Field)(nil)

func (iField Int64Field) Translate(value string, colIndex int) (interface{}, error) {
	if value == "" {
		return Int64Field{value: 0, colIndex: colIndex}, nil
	}
	res, err := strconv.ParseInt(value, 10, 64)
	if err != nil {
		return nil, err
	}
	return Int64Field{value: res, colIndex: colIndex}, nil
}

func (iField Int64Field) ColIndex() int {
	return iField.colIndex
}

func (iField Int64Field) GetStdValue() int64 {
	return iField.value
}

func (iField *Int64Field) SetValue(value int64) {
	iField.value = value
}

func (iField Int64Field) GetValue() interface{} {
	return iField.value
}

type StringField struct {
	value    string
	colIndex int
}

func NewStringField(value string) StringField {
	return StringField{value: value}
}

var _ Field = (*StringField)(nil)

func (sField StringField) Translate(value string, colIndex int) (interface{}, error) {
	return StringField{value: value, colIndex: colIndex}, nil
}

func (sField StringField) ColIndex() int {
	return sField.colIndex
}

func (sField StringField) GetStdValue() string {
	return sField.value
}

func (sField *StringField) SetValue(value string) {
	sField.value = value
}

func (sField StringField) GetValue() interface{} {
	return sField.value
}

type TimeField struct {
	value    time.Time
	colIndex int
}

func NewTimeField(value time.Time) TimeField {
	return TimeField{value: value}
}

var _ Field = (*TimeField)(nil)

func (tField TimeField) Translate(value string, colIndex int) (interface{}, error) {
	var (
		t   time.Time
		err error
	)
	if value == "" {
		return TimeField{value: t, colIndex: colIndex}, nil
	}
	t, err = time.ParseInLocation(_dateLayout, value, time.Local)
	if err != nil {
		return nil, err
	}
	return TimeField{value: t, colIndex: colIndex}, nil
}

func (tField TimeField) ColIndex() int {
	return tField.colIndex
}

func (tField TimeField) GetStdValue() time.Time {
	return tField.value
}

func (tField *TimeField) SetValue(value time.Time) {
	tField.value = value
}

func (tField TimeField) GetValue() interface{} {
	return tField.value
}

type FloatField struct {
	value    float64
	colIndex int
}

func NewFloatField(value float64) FloatField {
	return FloatField{value: value}
}

var _ Field = (*FloatField)(nil)

func (fField FloatField) Translate(value string, colIndex int) (interface{}, error) {
	if value == "" {
		return FloatField{0, colIndex}, nil
	}
	res, err := strconv.ParseFloat(value, 64)
	if err != nil {
		return nil, err
	}
	return FloatField{value: res, colIndex: colIndex}, nil
}

func (fField FloatField) ColIndex() int {
	return fField.colIndex
}

func (fField FloatField) GetStdValue() float64 {
	return fField.value
}

func (fField *FloatField) SetValue(value float64) {
	fField.value = value
}

func (fField FloatField) GetValue() interface{} {
	return fField.value
}

type BoolField struct {
	value    bool
	colIndex int
}

func NewBoolField(value bool) BoolField {
	return BoolField{value: value}
}

var _ Field = (*BoolField)(nil)

func (bField BoolField) Translate(value string, colIndex int) (interface{}, error) {
	if value == "" {
		return BoolField{false, colIndex}, nil
	}
	var res bool
	if value == "是" {
		res = true
	}
	return BoolField{res, colIndex}, nil
}

func (bField BoolField) ColIndex() int {
	return bField.colIndex
}

func (bField BoolField) GetStdValue() bool {
	return bField.value
}

func (bField *BoolField) SetValue(value bool) {
	bField.value = value
}

func (bField BoolField) GetValue() interface{} {
	return bField.value
}
