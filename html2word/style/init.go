package style

import "baliance.com/gooxml/document"

var Doc *document.Document = document.New()

const (
	// font
	NormalFontSize float64 = 14
	H1FontSize     float64 = 24
	H2FontSize     float64 = 22
	H3FontSize     float64 = 20
	H4FontSize     float64 = 18
	H5FontSize     float64 = 16

	// img
	A4Width  float64 = 210 * 2
	A4Height float64 = 297 * 2

	// color
	CodeBackGround     string = "#dfe3e7"
	HyperLinkFontColor string = "#0563C1"
)
