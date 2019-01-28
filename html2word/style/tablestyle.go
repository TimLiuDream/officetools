package style

// 合并格子的合并范围
type MergeCellScope struct {
	RowScope
	ColScope
}

// 行合并范围
type RowScope struct {
	Start int
	End   int
}

// 列合并范围
type ColScope struct {
	Start int
	End   int
}

func SetTable(cellMap map[string]string, mergeCellScopeMap map[string]*MergeCellScope) error {
	return nil
}
