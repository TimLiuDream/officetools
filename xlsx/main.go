package main

import (
	"fmt"
	"strconv"

	"github.com/tealeg/xlsx"
)

var (
	inFile  = "/Users/chain/Downloads/student1.xlsx"
	outFile = "/Users/tim/Downloads/out_student.xlsx"
	// outFile = "F:\\out_student.xlsx"
)

type Student struct {
	Name   string
	age    int
	Phone  string
	Gender string
	Mail   string
}

func Export() {
	style := xlsx.NewStyle()
	font := *xlsx.NewFont(20, "Verdana")
	border := *xlsx.NewBorder("thin", "thin", "thin", "thin")
	alignment := xlsx.Alignment{Vertical: "center", WrapText: true, ShrinkToFit: true}
	style.Font = font
	style.Border = border
	style.Alignment = alignment
	style.ApplyFont = true
	style.ApplyBorder = true
	style.ApplyAlignment = true

	file := xlsx.NewFile()
	sheet, err := file.AddSheet("sheet1") // sheet的名字不能是中文
	if err != nil {
		fmt.Printf(err.Error())
	}

	stus := getStudents()
	headRow := sheet.AddRow()

	aCell := headRow.AddCell()
	aCell.Value = "不会有乱码吧！！"
	aCell.SetStyle(style)

	//add data
	for _, stu := range stus {
		row := sheet.AddRow()
		nameCell := row.AddCell()
		nameCell.Value = stu.Name
		nameCell.SetStyle(style)
		ageCell := row.AddCell()
		ageCell.Value = strconv.Itoa(stu.age)
		ageCell.SetStyle(style)
		phoneCell := row.AddCell()
		phoneCell.Value = stu.Phone
		phoneCell.SetStyle(style)
		genderCell := row.AddCell()
		genderCell.Value = stu.Gender
		genderCell.SetStyle(style)
		mailCell := row.AddCell()
		mailCell.Value = stu.Mail
		mailCell.SetStyle(style)
	}
	err = sheet.SetColWidth(0, len(sheet.Cols)-1, 20)
	if err != nil {
		fmt.Printf(err.Error())
	}
	err = file.Save(outFile)
	if err != nil {
		fmt.Printf(err.Error())
	}
	fmt.Println("\n\nexport success")
}

func getStudents() []Student {
	students := make([]Student, 0)
	for i := 0; i < 10; i++ {
		stu := Student{}
		stu.Name = "name" + strconv.Itoa(i+1)
		stu.Mail = stu.Name + "@chairis.cn"
		stu.Phone = "1380013800\n" + strconv.Itoa(i)
		stu.age = 20
		stu.Gender = "男"
		students = append(students, stu)
	}
	return students
}

func main() {
	Export()
}
