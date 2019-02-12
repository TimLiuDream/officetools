package wordstyle

import (
	"log"
	"math"

	"baliance.com/gooxml/common"
	"baliance.com/gooxml/measurement"
)

// TODO 需要处理大图片的情况

// SetImage 往word写入图片
func SetImage(imgPath string, size string) error {
	img, err := common.ImageFromFile(imgPath)
	if err != nil {
		log.Fatalln(err)
		return err
	}
	imgRef, err := Doc.AddImage(img)
	if err != nil {
		log.Fatalln(err)
		return err
	}
	paragraph := Doc.AddParagraph()
	run := paragraph.AddRun()
	inline, err := run.AddDrawingInline(imgRef)
	if err != nil {
		log.Fatalln(err)
		return err
	}
	realX, realY := 0, 0
	if size == ImgSizeSmall {
		realX, realY = calculateRatioFit(img.Size.X, img.Size.Y, ImgSizeProportionSmall)
	} else if size == ImgSizeMedium {
		realX, realY = calculateRatioFit(img.Size.X, img.Size.Y, ImgSizeProportionMedium)
	} else if size == ImgSizeLarge {
		realX, realY = calculateRatioFit(img.Size.X, img.Size.Y, ImgSizeProportionLarge)
	}
	w := measurement.Distance(realX)
	h := measurement.Distance(realY)
	inline.SetSize(w, h)
	return nil
}

// calculateRatioFitSmall 计算图片缩放后的尺寸
func calculateRatioFit(srcWidth, srcHeight int, proportion float64) (int, int) {
	ratio := math.Min(proportion*A4Width/float64(srcWidth), proportion*A4Height/float64(srcHeight))
	return int(math.Ceil(float64(srcWidth) * ratio)), int(math.Ceil(float64(srcHeight) * ratio))
}
