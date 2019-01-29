package logger

import (
	"fmt"
	"io"
	"log"
	"os"
	"time"
)

var (
	// Trace 堆栈信息
	Trace *log.Logger
	// Info 提示信息
	Info *log.Logger
	// Warning 警告信息
	Warning *log.Logger
	// Error 错误信息
	Error *log.Logger
)

func init() {
	year := time.Now().Year()
	month := time.Now().Month()
	day := time.Now().Day()
	date := fmt.Sprintf("%d%d%d", year, month, day)
	file, err := os.OpenFile(date+"log.txt", os.O_CREATE|os.O_WRONLY|os.O_APPEND, 0666)
	if err != nil {
		log.Fatalln("Failed to open error log file:", err)
	}

	Trace = log.New(io.MultiWriter(file, os.Stdout), "TRACE:", log.Ldate|log.Ltime|log.Lshortfile)
	Info = log.New(io.MultiWriter(file, os.Stdout), "INFO:", log.Ldate|log.Ltime|log.Lshortfile)
	Warning = log.New(io.MultiWriter(file, os.Stdout), "WARNING:", log.Ldate|log.Ltime|log.Lshortfile)
	Error = log.New(io.MultiWriter(file, os.Stderr), "ERROR:", log.Ldate|log.Ltime|log.Lshortfile)
}
