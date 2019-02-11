package model

// TableCell 表格格子
type TableCell struct {
	RowIndex      int    // 行索引
	ColIndex      int    // 列索引
	IsVMerge      bool   // 是否是竖向合并
	IsVMergeStart bool   // 是否是竖向合并的开头格子
	HMerge        int    // 列合并数（横向合并数）
	Value         string // 格子的值
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
