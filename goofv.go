package goofv

import (
	"flag"
	"fmt"
	"log"
	"syscall"
	"unsafe"

	"golang.org/x/sys/windows/registry"
)

var (
	gkExcel, gkWord, gkPowerPoint                                           *syscall.DLL
	procExcelValidateFile, procWordValidateFile, procPowerPointValidateFile *syscall.Proc

	testFile = flag.String("file", "", "file to validate")
	mode     = flag.String("mode", "xls", "validation mode (must be one of: xls,doc,ppt)")
	format   = flag.String("format", "txt", "output format (must be one of: txt,csv,json,xml)")
	versions = getOfficeVersion()

	rootFolder string
)

func init() {
	forceFVLogging()
	gkExcel = syscall.MustLoadDLL(rootFolder + "GKExcel.dll")
	gkWord = syscall.MustLoadDLL(rootFolder + "GKWord.dll")
	gkPowerPoint = syscall.MustLoadDLL(rootFolder + "GKPowerPoint.dll")
	procExcelValidateFile = gkExcel.MustFindProc("FValidateExcelFile")
	procWordValidateFile = gkWord.MustFindProc("FValidateWordFile")
	procPowerPointValidateFile = gkPowerPoint.MustFindProc("FValidatePptFile")
}

func forceFVLogging() {
	for _, v := range versions {
		var key = `SOFTWARE\Microsoft\Office\` + v + `\`
		switch v {
		case "11.0":
		case "12.0":
		case "14.0":
			key += `Common\Security\FileValidation`
			rootFolder = "C:\\Program Files\\Microsoft Office\\Office14\\"
		case "15.0":
			key += `Common\Security\FileValidation`
			rootFolder = "C:\\Program Files\\Microsoft Office\\Office15\\"
		case "16.0":
			key += `Common\Security\FileValidation`
			rootFolder = "C:\\Program Files\\Microsoft Office\\root\\Office16\\"
		}

		k, _, e := registry.CreateKey(registry.CURRENT_USER, key, registry.ALL_ACCESS)
		if e != nil {
			log.Fatal("Could not open Logging Key: ", e)
		}
		defer k.Close()
		e = k.SetDWordValue("EnableLogging", 7)
		if e != nil {
			log.Fatal("Could not set Logging key: ", e)
		}
	}
}

func getOfficeVersion() []string {
	var ans []string
	k, err := registry.OpenKey(registry.CURRENT_USER, `SOFTWARE\Microsoft\Office`, registry.ALL_ACCESS)
	if err != nil {
		log.Fatal("Could not open Registry key: ", err)
	}
	defer k.Close()

	keys, err := k.ReadSubKeyNames(0)
	if err != nil {
		log.Fatal("Could not read subkeys: ", err)
	}
	for _, key := range keys {
		_, e := registry.OpenKey(k, key+`\Registration`, registry.QUERY_VALUE)
		if e == nil {
			ans = append(ans, key)
		}
	}
	return ans
}

func isValidFile(validator uintptr, file string) bool {
	filePtr := syscall.StringToUTF16Ptr(file)
	r1, _, er := syscall.Syscall(
		validator,
		2,
		uintptr(0),
		uintptr(unsafe.Pointer(filePtr)),
		0)
	if er != syscall.Errno(0) {
		log.Fatal(er)
	}
	return r1 != 0
}

func IsValidExcelFile(file string) bool {
	return isValidFile(procExcelValidateFile.Addr(), file)
}

func IsValidWordFile(file string) bool {
	return isValidFile(procWordValidateFile.Addr(), file)
}
func IsValidPowerPointFile(file string) bool {
	return isValidFile(procPowerPointValidateFile.Addr(), file)
}

func showHelp() {
	flag.Usage()
}

func showResults(file string, result bool) {
	switch *format {
	case "json":
		fmt.Printf(`{ "File": "%s", "IsValid": "%t" }`, file, result)
	case "csv":
		fmt.Printf(`"%s","%t"`, file, result)
	case "xml":
		fmt.Printf(`<File "isValid"="%t">%s</File>`, result, file)
	default:
		fmt.Printf("%s is Valid = %t\n", *testFile, result)

	}
}

func main() {
	flag.Parse()

	if *testFile == "" {
		fmt.Errorf("Must provide a file!")
		showHelp()
	} else {
		var result bool
		switch *mode {
		case "xls":
			result = IsValidExcelFile(*testFile)
		case "doc":
			result = IsValidWordFile(*testFile)
		case "ppt":
			result = IsValidPowerPointFile(*testFile)
		default:
			fmt.Println("Must be a valid mode!")
			showHelp()
		}
		showResults(*testFile, result)
	}
}
