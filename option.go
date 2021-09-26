package excel

type Option func(*Excel)

/**
OpenPassword set the excel open password
*/
func OpenPassword(password string) Option {
	return func(e *Excel) {
		e.password = password
	}
}

/**
SheetCount set the sheet count
*/
func SheetCount(sheetCount int) Option {
	return func(e *Excel) {
		e.sheetCount = sheetCount
	}
}

/**
SheetPrefix set the sheet prefix
*/
func SheetPrefix(sheetPrefix string) Option {
	return func(e *Excel) {
		e.sheetPrefix = sheetPrefix
	}
}

/**
ActiveSheet set the excel active sheet
*/
func ActiveSheet(sheetName ...string) Option {
	return func(e *Excel) {
		e.activeSheetNames = append(e.activeSheetNames, sheetName...)
	}
}

/**
HeaderRow set row for the header
*/
func HeaderRow(headerRow int) Option {
	return func(e *Excel) {
		e.headerRow = headerRow
	}
}
