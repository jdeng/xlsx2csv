package main

import (
	"fmt"
	"io/ioutil"
	"log"
	"os"
	"path"
	"path/filepath"
	"strings"

	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

func main() {
	ole.CoInitialize(0)

	ea, err := oleutil.CreateObject("Excel.Application")
	if err != nil {
		log.Fatal(err)
	}
	defer ea.Release()

	excel, _ := ea.QueryInterface(ole.IID_IDispatch)
	defer excel.Release()

	log.Printf("Excel Version: %s\n", oleutil.MustGetProperty(excel, "Version").ToString())

	//	oleutil.PutProperty(excel, "Visible", true) // hide window
	oleutil.PutProperty(excel, "DisplayAlerts", false)
	oleutil.PutProperty(excel, "ScreenUpdating", true)

	target := os.Args[1]
	fi, err := os.Stat(target)
	if err != nil {
		log.Fatal(err)
	}

	if fi.IsDir() {
		files, _ := ioutil.ReadDir(target)
		for i, file := range files {
			file, _ := filepath.Abs(filepath.Join(target, file.Name()))
			log.Printf("Processing #%d %s\n", i, file)
			saveWorkbook(excel, file, "out")
		}
	} else {
		saveWorkbook(excel, target, ".")
	}

	oleutil.MustCallMethod(excel, "Quit")
}

func saveWorkbook(excel *ole.IDispatch, file, dir string) {
	workbooks := oleutil.MustGetProperty(excel, "WorkBooks").ToIDispatch()
	defer workbooks.Release()

	books := oleutil.MustCallMethod(workbooks, "open", file, nil, true).ToIDispatch()
	defer books.Release()

	sheets := oleutil.MustGetProperty(excel, "Sheets").ToIDispatch()
	n := int(oleutil.MustGetProperty(sheets, "Count").Val)
	sheets.Release()

	log.Printf("Total %d sheet(s)\n", n)

	name := filepath.Base(file)
	name = strings.TrimSuffix(name, path.Ext(name))

	const xlCSVUTF8 = 62
	for i := 1; i <= n; i++ {
		worksheets := oleutil.MustGetProperty(excel, "Worksheets", i).ToIDispatch()
		oleutil.MustCallMethod(worksheets, "Select")

		fp := filepath.Join(dir, fmt.Sprintf("%s-%d.csv", name, i))
		fp, _ = filepath.Abs(fp)
		log.Printf("Saving sheet#%d to %s\n", i, fp)
		activeWorkBook := oleutil.MustGetProperty(excel, "ActiveWorkBook").ToIDispatch()
		oleutil.MustCallMethod(activeWorkBook, "SaveAs", fp, xlCSVUTF8, nil, nil)

		activeWorkBook.Release()
		worksheets.Release()
	}

	oleutil.MustCallMethod(workbooks, "Close")
}
