package model

// 合并格子的合并范围
type MergeCellScope struct {
	VMerge int // 行合并（竖向合并数）
	HMerge int // 列合并（横向合并数）
	Value  string
}
