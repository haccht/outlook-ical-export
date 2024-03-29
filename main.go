package main

import (
	"flag"
	"os"
	"path/filepath"
	"time"

	ole "github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

var (
	duration = flag.Int("d", 14, "the duration in days that the events should extracted")
	filename = flag.String("o", "calendar.ics", "the ical output filename")
)

func main() {
	flag.Parse()

	ole.CoInitialize(0)
	unknown, _ := oleutil.CreateObject("Outlook.Application")
	outlook, _ := unknown.QueryInterface(ole.IID_IDispatch)

	ns := oleutil.MustCallMethod(outlook, "GetNamespace", "MAPI").ToIDispatch()
	folder := oleutil.MustCallMethod(ns, "GetDefaultFolder", 9).ToIDispatch()
	export := oleutil.MustCallMethod(folder, "GetCalendarExporter").ToIDispatch()

	cwd, _ := os.Getwd()
	outfile := filepath.Join(cwd, *filename)
	enddate := time.Now().Add(time.Duration(*duration*24) * time.Hour)

	oleutil.MustPutProperty(export, "CalendarDetail", 2)
	oleutil.MustPutProperty(export, "EndDate", enddate)
	oleutil.MustPutProperty(export, "IncludeWholeCalendar", false)
	oleutil.MustPutProperty(export, "IncludeAttachments", false)
	oleutil.MustPutProperty(export, "IncludePrivateDetails", true)
	oleutil.MustPutProperty(export, "RestrictToWorkingHours", false)

	oleutil.MustCallMethod(export, "SaveAsICal", outfile)
}
