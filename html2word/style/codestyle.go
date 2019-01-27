package style

import (
	"baliance.com/gooxml/color"
	"baliance.com/gooxml/schema/soo/wml"
	"strings"
)

func SetCode(code string) {
	paragraph := Doc.AddParagraph()
	run := paragraph.AddRun()

	run.Properties().SetSize(20)
	run.Properties().SetColor(color.Black)
	run.Properties().SetHighlight(wml.ST_HighlightColorLightGray)

	for _, c := range splitCode(code) {
		run.AddTab()
		run.AddText(c)
	}

	run.AddBreak()
}

func splitCode(code string) []string {
	return strings.Split(code, "\n")
}
