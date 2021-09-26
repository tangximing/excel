package excel

import (
	"fmt"
	"strings"
)

type CellError struct {
	rowIndex, colIndex int
	paths              []string
	err                error
}

func (e *CellError) Error() string {
	return fmt.Sprintf("单元格填写错误。表头：%s, 行列：(%d, %d)",
		strings.Join(e.paths, _tagPathSplitter), e.rowIndex, e.colIndex)
}
