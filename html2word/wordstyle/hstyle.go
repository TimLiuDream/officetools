package wordstyle

import (
	"baliance.com/gooxml/measurement"
)

// SetH 往word中写入标题
func SetH(text string, tag string) {
	var size measurement.Distance
	switch tag {
	case "h1":
		size = measurement.Distance(H1FontSize)
	case "h2":
		size = measurement.Distance(H2FontSize)
	case "h3":
		size = measurement.Distance(H3FontSize)
	case "h4":
		size = measurement.Distance(H4FontSize)
	case "h5":
		size = measurement.Distance(H5FontSize)
	}
	paragraph := Doc.AddParagraph()
	run := paragraph.AddRun()
	run.Properties().SetBold(true)
	run.Properties().SetSize(size)
	run.AddText(text)
	run.AddBreak()
}
