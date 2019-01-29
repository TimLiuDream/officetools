package model

// TableCell 表格格子
type TableCell struct {
	RowIndex int // 行索引
	ColIndex int // 列索引
	VMerge   int // 行合并（竖向合并数）
	HMerge   int // 列合并（横向合并数）
	Value    string
}
