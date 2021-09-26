package excel

import (
	"fmt"
	"reflect"
	"runtime"
	"strings"
	"sync"

	"github.com/360EntSecGroup-Skylar/excelize/v2"
	"github.com/panjf2000/ants/v2"
	"github.com/pkg/errors"
)

type Importer struct {
	// value of current cell node
	value string
	// beginning col index of current node cell
	colIndexStart int
	// end col index of current node cell
	colIndexEnd int
	// beginning row index of current node cell
	rowIndexStart int
	// end row index of current node cell
	rowIndexEnd int

	// path store the path from root to current node
	path []string
	// children nodes of current node
	childImporters []*Importer
}

/**
GetColIndexPos return the col start index and col end index of root cell
*/
func (root *Importer) GetColIndexPos() (colIndexStart, colIndexEnd int) {
	return root.colIndexStart, root.colIndexEnd
}

/**
GetColIndexPos return the row start index and row end index of root cell
*/
func (root *Importer) GetRowIndexPos() (rowIndexStart, rowIndexEnd int) {
	return root.rowIndexStart, root.rowIndexEnd
}

/**
SubImporter return the sub importer by excel path
*/
func (root *Importer) SubImporter(path string) *Importer {
	if path == "" {
		return nil
	}

	paths := strings.Split(path, _tagPathSplitter)
	if len(root.path) > 0 {
		return root.subImporter(append([]string{root.value}, paths...))
	}
	return root.subImporter(paths)
}

func (root *Importer) subImporter(path []string) *Importer {
	if len(path) == 0 {
		return nil
	}
	if len(path) == 1 {
		if path[0] == root.value {
			return root
		}
		return nil
	}
	// if root is the root of the excel tree, it's a fake node, it's path is nil, so handle it special,
	// no need to consume the path param, just go on.
	if root.path == nil || root.value == path[0] {
		if root.path != nil {
			path = path[1:]
		}
		for _, node := range root.childImporters {
			if importer := node.subImporter(path); importer != nil {
				return importer
			}
		}
	}
	return nil
}

/**
ScanRow scan an excel row to structs. The func also support to scan a row by relative path
ex: if a leaf node's path is `excel:"a|b|c"`, we can define a struct field `test` which has a tag `excel:"b|c"`, and
it can scan because the path for the field match the leaf node's behind path
Note: responses must be struct pointer types
*/
func (root *Importer) ScanRow(row []string, responses ...interface{}) (err error) {
	defer func() {
		if p := recover(); p != nil {
			err = fmt.Errorf("internal error: %v", p)
		}
	}()

	leafNodes := root.getLeafNodes()
	leafNodesLength := len(leafNodes)
	rowLength := len(row)
	if leafNodesLength < rowLength {
		row = row[rowLength-leafNodesLength:]
	}
	if rowLength < leafNodesLength {
		for i := 0; i < leafNodesLength-rowLength; i++ {
			row = append(row, "")
		}
	}

	for _, resp := range responses {
		v := reflect.ValueOf(resp).Elem()
		for i := 0; i < reflect.Indirect(v).NumField(); i++ {
			field := reflect.Indirect(v).Type().Field(i)
			tag := field.Tag.Get(_tagFlag)
			path := strings.Split(tag, _tagPathSplitter)
			for j, leafNode := range leafNodes {
				nodePath := leafNode.path[len(leafNode.path)-len(path):]
				if reflect.DeepEqual(nodePath, path) {
					var setValue interface{}
					setValue, err = reflect.Indirect(v).Field(i).Interface().(Field).Translate(row[j], leafNode.colIndexStart)
					if err != nil {
						err = &CellError{rowIndex: j, colIndex: i, err: err}
						return
					}

					reflect.Indirect(v).Field(i).Set(reflect.ValueOf(setValue))
					break
				}
			}
		}
	}
	return
}

func (root *Importer) getLeafNodes() []*Importer {
	if root == nil {
		return nil
	}

	var res []*Importer
	if len(root.childImporters) == 0 {
		res = append(res, root)
		return res
	}
	for _, im := range root.childImporters {
		res = append(res, im.getLeafNodes()...)
	}
	return res
}

type AsyncScanExRes struct {
	Responses []interface{}
	Err       error
}

/**
AsyncScanRows scan rows to responses async
*/
func (root *Importer) AsyncScanRows(rows [][]string, responses ...interface{}) chan *AsyncScanExRes {
	pool, _ := ants.NewPool(runtime.NumCPU())

	ch := make(chan *AsyncScanExRes, len(rows))
	var wg sync.WaitGroup
	for i := range rows {
		wg.Add(1)

		index := i
		_ = pool.Submit(func() {
			defer wg.Done()

			// we need to make a copy of the receiver for the row data
			var respParams []interface{}
			for _, resp := range responses {
				respParams = append(respParams, reflect.New(reflect.Indirect(reflect.ValueOf(resp).Elem()).Type()).Interface())
			}

			err := root.ScanRow(rows[index], respParams...)
			ch <- &AsyncScanExRes{Responses: respParams, Err: err}
		})
	}

	go func(wg *sync.WaitGroup, ch chan *AsyncScanExRes, wp *ants.Pool) {
		wg.Wait()
		close(ch)
		wp.Release()
	}(&wg, ch, pool)

	return ch
}

func (root *Importer) IsHeaderConsistent(responses ...interface{}) (isConsistent bool, err error) {
	defer func() {
		if p := recover(); p != nil {
			err = fmt.Errorf("internal error: %v", p)
		}
	}()

	leafNodes := root.getLeafNodes()
	lastRespLength := 0
	for _, resp := range responses {
		t := reflect.TypeOf(resp)
		switch t.Kind() {
		case reflect.Ptr:
			t = t.Elem()
		default:
			err = errors.New("response is not ptr type")
			return
		}

		fieldNum := t.NumField()
		for i := 0; i < fieldNum; i++ {
			field := t.Field(i)
			tag := field.Tag.Get(_tagFlag)
			path := strings.Split(tag, _tagPathSplitter)

			leafNode := leafNodes[lastRespLength+i]
			if !reflect.DeepEqual(leafNode.path, path) {
				return
			}
		}

		lastRespLength += fieldNum
	}

	if lastRespLength != len(leafNodes) {
		return
	}

	isConsistent = true
	return
}

/**
getChildrenColBoundary return start, end index
start - the node's index in mergeCells which is root's first child
end -  the node's index in mergeCells which is root's last child
*/
func (root *Importer) getChildrenColBoundary(mergeCells []excelize.MergeCell) (start, end int, err error) {
	var (
		// the beginning col index of current cell
		colIndexStart int
		// the end col index of current cell
		colIndexEnd int
		// current cell's row index
		rowIndex int
	)

	// set to negative to make sure it also work when there is no child.
	end = -1

	for i, cell := range mergeCells {
		startAxis, endAxis := cell.GetStartAxis(), cell.GetEndAxis()
		if colIndexStart, rowIndex, err = excelize.CellNameToCoordinates(startAxis); err != nil {
			err = errors.Wrap(err, "excelize.CellNameToCoordinates")
			return
		}
		if colIndexEnd, _, err = excelize.CellNameToCoordinates(endAxis); err != nil {
			err = errors.Wrap(err, "excelize.CellNameToCoordinates")
			return
		}
		if rowIndex > root.rowIndexEnd && colIndexStart >= root.colIndexStart && colIndexEnd <= root.colIndexEnd {
			if colIndexStart == root.colIndexStart {
				start = i
			}
			if colIndexEnd == root.colIndexEnd {
				end = i
				break
			}
		}
	}

	return
}

/**
getRowsBeginIndex return the beginning row index of excel (except of mergeCell headers)
*/
func (root *Importer) getRowsBeginIndex() int {
	if root == nil {
		return 0
	}

	im := root
	for len(im.childImporters) != 0 {
		im = im.childImporters[0]
	}

	return im.rowIndexEnd
}

/**
buildNode build a node from a merge cell
*/
func buildNode(mergeCell excelize.MergeCell) (importer *Importer, err error) {
	importer = new(Importer)
	importer.value = mergeCell.GetCellValue()
	startAxis, endAxis := mergeCell.GetStartAxis(), mergeCell.GetEndAxis()
	if importer.colIndexStart, importer.rowIndexStart, err = excelize.CellNameToCoordinates(startAxis); err != nil {
		err = errors.Wrap(err, "excelize.CellNameToCoordinates")
		return
	}
	if importer.colIndexEnd, importer.rowIndexEnd, err = excelize.CellNameToCoordinates(endAxis); err != nil {
		err = errors.Wrap(err, "excelize.CellNameToCoordinates")
		return nil, err
	}

	return
}

/**
buildChildNodes build child nodes by mergeCells
*/
func buildChildNodes(root *Importer, mergeCells []excelize.MergeCell) (children []*Importer, err error) {
	if root == nil {
		return
	}

	// get the root's first and last child in mergeCells
	start, end, err := root.getChildrenColBoundary(mergeCells)
	if err != nil {
		err = errors.Wrap(err, "root.getChildrenColBoundary")
		return
	}
	if end < 0 {
		return
	}

	for i := start; i <= end; i++ {
		node, e := buildNode(mergeCells[i])
		if e != nil {
			err = errors.Wrap(e, "buildNodeByCell")
			return
		}

		// children's path
		node.path = append(node.path, root.path...)
		node.path = append(node.path, mergeCells[i].GetCellValue())
		if node.childImporters, err = buildChildNodes(node, mergeCells); err != nil {
			err = errors.Wrap(err, "buildChildNodes")
			return
		}

		children = append(children, node)
	}

	return
}
