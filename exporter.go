package excel

import (
	"reflect"
	"strings"

	"github.com/360EntSecGroup-Skylar/excelize/v2"
	"github.com/pkg/errors"
)

type initData func(rows []interface{}) (err error)

func (e *Excel) initFromData(rows []interface{}) (err error) {
	if len(rows) == 0 {
		return
	}

	// parse header
	header, err := parseHeader(rows[0])
	if err != nil {
		err = errors.Wrap(err, "parseHeader")
		return
	}
	_, err = e.writeHeader(header, 1, 0)
	if err != nil {
		err = errors.Wrap(err, "e.writeHeader")
		return
	}
	dataRow := header.getHeight()
	if err = e.writeData(rows, dataRow); err != nil {
		err = errors.Wrap(err, "e.writeData")
		return
	}

	return
}

func (e *Excel) writeHeader(header *header, col, row int) (span int, err error) {
	if header == nil {
		return
	}

	var childrenSpan int
	for _, child := range header.children {
		var childSpan int
		childSpan, err = e.writeHeader(child, col+childrenSpan, row+1)
		if err != nil {
			err = errors.Wrap(err, "e.writeHeader")
			return
		}

		childrenSpan += childSpan
	}
	if len(header.children) == 0 {
		span = 1
	} else {
		span = childrenSpan

		if !header.isDummy {
			// merge cells
			var hCell, vCell string
			hCell, err = excelize.CoordinatesToCellName(col, row)
			if err != nil {
				err = errors.Wrap(err, "excelize.CoordinatesToCellName")
				return
			}
			vCell, err = excelize.CoordinatesToCellName(col+childrenSpan-1, row)
			if err != nil {
				err = errors.Wrap(err, "excelize.CoordinatesToCellName")
				return
			}
			for _, sheet := range e.activeSheetNames {
				err = e.ex.MergeCell(sheet, hCell, vCell)
				if err != nil {
					err = errors.Wrap(err, "e.ex.MergeCell")
					return
				}

				err = e.ex.SetCellStyle(sheet, hCell, vCell, e.fieldStyleId)
				if err != nil {
					err = errors.Wrap(err, "e.ex.SetCellStyle")
					return
				}
			}
		}
	}

	if header.isDummy {
		// fake node
		return
	}

	axis, err := excelize.CoordinatesToCellName(col, row)
	if err != nil {
		err = errors.Wrap(err, "excelize.CoordinatesToCellName")
		return
	}
	for _, sheet := range e.activeSheetNames {
		err = e.ex.SetCellValue(sheet, axis, header.title)
		if err != nil {
			err = errors.Wrap(err, "e.ex.SetCellValue")
			return
		}
	}

	return
}

type header struct {
	isDummy  bool // the root is fake node
	parent   string
	title    string
	children []*header
}

func (h header) getHeight() (height int) {
	if len(h.children) == 0 {
		return 1
	}

	height = h.children[0].getHeight() + 1
	return
}

func parseHeader(row interface{}) (h *header, err error) {
	h = &header{isDummy: true}

	var paths [][]string
	v := reflect.ValueOf(row).Elem()
	for i := 0; i < reflect.Indirect(v).NumField(); i++ {
		field := reflect.Indirect(v).Type().Field(i)
		tag := field.Tag.Get(_tagFlag)
		path := strings.Split(tag, _tagPathSplitter)
		paths = append(paths, path)
	}
	if len(paths) == 0 {
		return
	}

	pathDepth := len(paths[0])
	for _, path := range paths {
		if len(path) != pathDepth {
			err = errors.New("path depth is not same")
			return
		}
	}

	h.children = getHeadersFromPaths(paths, 0, pathDepth)
	return
}

func getHeadersFromPaths(paths [][]string, colIdx, colMax int) (hs []*header) {
	if colIdx == colMax {
		return
	}

	hMap := make(map[string]*header)
	for _, path := range paths {
		title := path[colIdx]
		var parent string
		if colIdx > 0 {
			parent = path[colIdx-1]
		}
		h, ok := hMap[title]
		if !ok {
			h = &header{parent: parent, title: title}
			hMap[title] = h

			// keep the title order
			hs = append(hs, h)
		}
	}

	childColIdx := colIdx + 1
	childHs := getHeadersFromPaths(paths, childColIdx, colMax)
	for _, childH := range childHs {
		parentH := hMap[childH.parent]
		parentH.children = append(parentH.children, childH)
	}

	return hs
}

func (e *Excel) writeData(rows []interface{}, rowStart int) (err error) {
	l := len(rows)
	if l == 0 {
		return
	}

	if l < e.sheetCount {
		err = errors.New("data rows is smaller than sheet count")
		return
	}

	sheetRowSize := l / e.sheetCount
	for idx := 0; idx < e.sheetCount; idx++ {
		var sheetRows []interface{}
		if idx == e.sheetCount-1 {
			sheetRows = rows[idx*sheetRowSize:]
		} else {
			sheetRows = rows[idx*sheetRowSize : (idx+1)*sheetRowSize]
		}

		sheetRowStart := rowStart
		sheet := e.activeSheetNames[idx]
		for _, row := range sheetRows {
			v := reflect.ValueOf(row).Elem()
			for i := 0; i < reflect.Indirect(v).NumField(); i++ {
				var axis string
				axis, err = excelize.CoordinatesToCellName(i+1, sheetRowStart)
				if err != nil {
					err = errors.Wrap(err, "excelize.CoordinatesToCellName")
					return
				}

				var value interface{}
				field := reflect.Indirect(v).Field(i)
				if importField, ok := field.Interface().(Field); ok {
					value = importField.GetValue()
				} else {
					value = field
				}
				err = e.ex.SetCellValue(sheet, axis, value)
				if err != nil {
					err = errors.Wrap(err, "e.ex.SetCellValue")
					return
				}
			}

			sheetRowStart++
		}
	}

	return
}
