package wordstyle

import (
	"baliance.com/gooxml/color"
	"baliance.com/gooxml/measurement"
	"baliance.com/gooxml/schema/soo/wml"
)

// SetHyperlink 往word写超链接
func SetHyperlink(link string) {
	hlStyle := Doc.Styles.AddStyle("Hyperlink", wml.ST_StyleTypeCharacter, false)
	hlStyle.SetName("Hyperlink")
	hlStyle.SetBasedOn("DefaultParagraphFont")
	hlStyle.RunProperties().Color().SetThemeColor(wml.ST_ThemeColorHyperlink)
	clr := color.FromHex(HyperLinkFontColor)
	hlStyle.RunProperties().Color().SetColor(clr)
	hlStyle.RunProperties().SetUnderline(wml.ST_UnderlineSingle, clr)

	paragraph := Doc.AddParagraph()
	hyperlink := paragraph.AddHyperLink()
	hyperlink.SetTarget(link)
	run := hyperlink.AddRun()
	run.Properties().SetStyle("Hyperlink")
	run.AddText(link)
	run.AddBreak()
}

// SetP 往word写段落
func SetP(text string) {
	paragraph := Doc.AddParagraph()
	run := paragraph.AddRun()
	run.Properties().SetSize(measurement.Distance(NormalFontSize))
	run.AddText(text)
	run.AddBreak()
}
