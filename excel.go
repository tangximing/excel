package excel

import (
	"fmt"
	"io"

	"github.com/360EntSecGroup-Skylar/excelize/v2"
	"github.com/pkg/errors"
)

type Excel struct {
	ex *excelize.File

	password    string
	sheetCount  int
	sheetPrefix string
	headerRow   int

	importers           []*Importer
	activeSheetNames    []string
	asyncScanWorkerNums int
	humanErrorMsg       bool

	// style
	fieldStyleId int
}

func NewExcelFromFile(file string, options ...Option) (e *Excel, err error) {
	e = newExcel()
	for _, option := range options {
		option(e)
	}
	if e.ex, err = excelize.OpenFile(file, excelize.Options{Password: e.password}); err != nil {
		err = errors.Wrap(err, "excelize.OpenFile")
		return
	}

	if err = e.postInitialize(nil, nil); err != nil {
		err = errors.Wrapf(err, "s.postInitialize")
		return
	}

	return
}

func NewExcelFromReader(reader io.Reader, options ...Option) (e *Excel, err error) {
	e = newExcel()
	for _, option := range options {
		option(e)
	}

	if e.ex, err = excelize.OpenReader(reader, excelize.Options{Password: e.password}); err != nil {
		err = errors.Wrap(err, "excelize.OpenReader")
		return
	}

	if err = e.postInitialize(nil, nil); err != nil {
		err = errors.Wrapf(err, "s.postInitialize")
		return
	}

	return
}

func NewExcelFromData(rows []interface{}, options ...Option) (e *Excel, err error) {
	e = newExcel()
	for _, option := range options {
		option(e)
	}

	e.ex = excelize.NewFile()
	for i := 1; i <= e.sheetCount; i++ {
		sheetName := fmt.Sprintf("%s%d", e.sheetPrefix, i)
		e.ex.NewSheet(sheetName)
	}

	if e.sheetPrefix != _defaultSheetPrefix {
		// delete default sheet
		e.ex.DeleteSheet("Sheet1")
	}

	if err = e.postInitialize(rows, e.initFromData); err != nil {
		err = errors.Wrapf(err, "s.postInitialize")
		return
	}

	return
}

func newExcel() (e *Excel) {
	e = new(Excel)
	e.sheetCount = 1
	e.sheetPrefix = _defaultSheetPrefix

	return e
}

func (e *Excel) postInitialize(rows []interface{}, initData initData) error {
	if len(e.activeSheetNames) != 0 {
		for _, activeSheetName := range e.activeSheetNames {
			activeSheetIndex := e.ex.GetSheetIndex(activeSheetName)
			if activeSheetIndex == -1 {
				return errors.Errorf("sheet name %s is invalid or doesn't exist", activeSheetName)
			}
		}
	} else {
		e.activeSheetNames = e.ex.GetSheetList()
	}

	if len(e.activeSheetNames) == 0 {
		return errors.New("no sheet exist")
	}

	// now just set one sheet active
	e.ex.SetActiveSheet(e.ex.GetSheetIndex(e.activeSheetNames[_defaultSheetIndex]))

	// set style
	if err := e.initStyle(); err != nil {
		return fmt.Errorf("init excel style error:(%+v)", err)
	}

	if len(rows) != 0 && initData != nil {
		err := initData(rows)
		if err != nil {
			return errors.Wrap(err, "initData")
		}
	}

	if err := e.initImporters(); err != nil {
		return fmt.Errorf("init excel importers error:(%+v)", err)
	}

	return nil
}

func (e *Excel) initStyle() (err error) {
	alignment := &excelize.Alignment{
		Horizontal:  "center",
		Vertical:    "center",
		ShrinkToFit: true,
	}

	e.fieldStyleId, err = e.ex.NewStyle(&excelize.Style{
		Alignment: alignment,
	})
	if err != nil {
		return
	}

	return
}

func (e *Excel) initImporters() (err error) {
	for _, sheetName := range e.activeSheetNames {
		root := new(Importer)
		root.value = sheetName
		root.colIndexStart = _defaultColStart
		if root.colIndexEnd, err = e.getSheetLastColIndex(sheetName); err != nil {
			err = errors.Wrapf(err, "e.getSheetLastColIndex")
			return
		}

		// get sheet headers in merge cells format
		var mergeCells []excelize.MergeCell
		mergeCells, err = e.getHeaders(sheetName)
		if err != nil {
			err = errors.Wrapf(err, "e.getHeaders")
			return
		}

		if root.childImporters, err = buildChildNodes(root, mergeCells); err != nil {
			return
		}

		e.importers = append(e.importers, root)
	}

	return
}

func (e *Excel) getSheetLastColIndex(sheet string) (int, error) {
	cols, err := e.ex.GetCols(sheet)
	if err != nil {
		return 0, err
	}
	return len(cols), nil
}

func (e *Excel) getHeaders(sheet string) (headers []excelize.MergeCell, err error) {
	if e.headerRow == 0 {
		headers, err = e.ex.GetMergeCells(sheet)
	} else {
		headers, err = e.getHeadersFromRow(sheet)
	}

	return
}

func (e *Excel) getHeadersFromRow(sheet string) (headers []excelize.MergeCell, err error) {
	headerRows, err := e.getHeaderRows(sheet)
	if err != nil {
		err = errors.Wrap(err, "e.getHeaderRows")
		return
	}
	if len(headerRows) == 0 {
		return
	}

	for i, row := range headerRows {
		if len(row) == 0 || row[0] == "" {
			continue
		}

		type headerIndex struct {
			header     string
			start, end int
		}
		headerIndices := make([]headerIndex, 0)
		for j, col := range row {
			if col == "" {
				continue
			}

			l := len(headerIndices)
			if l > 0 {
				headerIndices[l-1].end = j - 1
			}
			headerIndices = append(headerIndices, headerIndex{
				header: col,
				start:  j,
			})
		}
		headerIndices[len(headerIndices)-1].end = len(row) - 1

		for _, headerIndex := range headerIndices {
			var header excelize.MergeCell
			header, err = e.getMergeCell(headerIndex.start+1, headerIndex.end+1, i+1, headerIndex.header)
			if err != nil {
				err = errors.Wrap(err, "e.getMergeCell")
				return
			}
			headers = append(headers, header)
		}

	}
	return
}

func (e *Excel) getHeaderRows(sheet string) ([][]string, error) {
	rows, err := e.ex.Rows(sheet)
	if err != nil {
		return nil, err
	}
	results := make([][]string, 0, 64)

	headerRow := e.headerRow
	for rows.Next() && headerRow > 0 {
		row, err := rows.Columns()
		if err != nil {
			break
		}
		results = append(results, row)
		headerRow--
	}
	return results, nil
}

func (e *Excel) getMergeCell(startCol, endCol, row int, value string) (mergeCell excelize.MergeCell, err error) {
	var startAxis, endAxis string
	startAxis, err = excelize.CoordinatesToCellName(startCol, row)
	if err != nil {
		err = errors.Wrap(err, "excelize.CoordinatesToCellName")
		return
	}
	endAxis, err = excelize.CoordinatesToCellName(endCol, row)
	if err != nil {
		err = errors.Wrap(err, "excelize.CoordinatesToCellName")
		return
	}

	mergeCell = make([]string, 2)
	mergeCell[0] = fmt.Sprintf("%s:%s", startAxis, endAxis)
	mergeCell[1] = value
	return
}
