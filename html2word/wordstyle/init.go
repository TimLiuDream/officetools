package wordstyle

import "baliance.com/gooxml/document"

// Doc 文档对象
var Doc = document.New()

const (
	NormalFontSize float64 = 14
	H1FontSize     float64 = 24
	H2FontSize     float64 = 22
	H3FontSize     float64 = 20
	H4FontSize     float64 = 18
	H5FontSize     float64 = 16

	A4Width  float64 = 210 * 2
	A4Height float64 = 297 * 2

	CodeBackGround     string = "#dfe3e7"
	HyperLinkFontColor string = "#0563C1"
)
