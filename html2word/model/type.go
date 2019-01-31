package model

// TableCell 表格格子
type TableCell struct {
	RowIndex int // 行索引
	ColIndex int // 列索引
	VMerge   int // 行合并数（竖向合并数）
	HMerge   int // 列合并数（横向合并数）
	Value    string
}

// TableRowTitle 表格行标题
type TableRowTitle struct {
	ColIndex int    //列索引
	Title    string //标题
}

// TableColTitle 表格列标题
type TableColTitle struct {
	RowIndex int    //行索引
	Title    string //标题
}
