package style

import (
	"baliance.com/gooxml/common"
	"baliance.com/gooxml/measurement"
	"math"
)

func SetImage(imgPath string) error {
	img, err := common.ImageFromFile(imgPath)
	if err != nil {
		return err
	}
	imgRef, err := Doc.AddImage(img)
	if err != nil {
		return err
	}
	paragraph := Doc.AddParagraph()
	run := paragraph.AddRun()
	inline, err := run.AddDrawingInline(imgRef)
	if err != nil {
		return err
	}
	realX, realY := calculateRatioFit(img.Size.X, img.Size.Y)
	w := measurement.Distance(realX)
	h := measurement.Distance(realY)
	inline.SetSize(w, h)
	return nil
}

// 计算图片缩放后的尺寸
func calculateRatioFit(srcWidth, srcHeight int) (int, int) {
	ratio := math.Min(A4Width/float64(srcWidth), A4Height/float64(srcHeight))
	return int(math.Ceil(float64(srcWidth) * ratio)), int(math.Ceil(float64(srcHeight) * ratio))
}
