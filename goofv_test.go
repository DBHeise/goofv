package goofv

import (
	"path"
	"runtime"
	"strings"
	"testing"
)

var (
	excelGood = "testfiles/test.good.xls"
	excelBad  = "testfiles/test.corrupt.xls"
	wordGood  = "testfiles/test.good.doc"
	wordBad   = "testfiles/test.corrupt.doc"
	pptGood   = "testfiles/test.good.ppt"
	pptBad    = "testfiles/test.corrupt.ppt"
)

func init() {
	_, me, _, _ := runtime.Caller(1)
	baseFolder := path.Dir(me)
	excelGood = strings.Replace(path.Join(baseFolder, excelGood), "/", "\\", -1)
	excelBad = strings.Replace(path.Join(baseFolder, excelBad), "/", "\\", -1)

	wordGood = strings.Replace(path.Join(baseFolder, wordGood), "/", "\\", -1)
	wordBad = strings.Replace(path.Join(baseFolder, wordBad), "/", "\\", -1)

	pptGood = strings.Replace(path.Join(baseFolder, pptGood), "/", "\\", -1)
	pptBad = strings.Replace(path.Join(baseFolder, pptBad), "/", "\\", -1)
}
func TestExcelGood(t *testing.T) {
	result := IsValidExcelFile(excelGood)
	if !result {
		t.Fail()
	}
}
func TestExcelBad(t *testing.T) {
	result := IsValidExcelFile(excelBad)
	if result {
		t.Fail()
	}
}
func TestWordGood(t *testing.T) {
	result := IsValidWordFile(wordGood)
	if !result {
		t.Fail()
	}
}
func TestWordBad(t *testing.T) {
	result := IsValidWordFile(wordBad)
	if result {
		t.Fail()
	}
}
func TestPPTGood(t *testing.T) {
	result := IsValidPowerPointFile(pptGood)
	if !result {
		t.Fail()
	}
}
func TestPPTBad(t *testing.T) {
	result := IsValidPowerPointFile(pptBad)
	if result {
		t.Fail()
	}
}
