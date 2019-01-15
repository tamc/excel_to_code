// Compiled version of /Users/tamc/Documents/github/excel_to_code/spec/test.xlsx
package excelspreadsheet

import (
	"fmt"
	"strconv"
)

type spreadsheet struct {
  valuesA1 cachedValue
  valuesB1 cachedValue
  valuesA2 cachedValue
  valuesB2 cachedValue
  valuesA3 cachedValue
  valuesB3 cachedValue
  valuesC3 cachedValue
  valuesA4 cachedValue
  valuesB4 cachedValue
  valuesA5 cachedValue
  valuesB5 cachedValue
  valuesA6 cachedValue
  valuesB6 cachedValue
  valuesA7 cachedValue
  valuesB7 cachedValue
  valuesA8 cachedValue
  valuesB8 cachedValue
  valuesA9 cachedValue
  valuesB9 cachedValue
  valuesA10 cachedValue
  valuesB10 cachedValue
  valuesA11 cachedValue
  valuesB11 cachedValue
  valuesA12 cachedValue
  valuesB12 cachedValue
  valuesC1 cachedValue
}

func New() spreadsheet {
  return spreadsheet{}
}

func (s *spreadsheet) ValuesA1() (interface{}, error) {
  if !s.valuesA1.isCached() {
    s.valuesA1.set("String")
  }
  return s.valuesA1.get()
}

func (s *spreadsheet) ValuesB1() (interface{}, error) {
  if !s.valuesB1.isCached() {
    s.valuesB1.set("String")
  }
  return s.valuesB1.get()
}

func (s *spreadsheet) ValuesA2() (interface{}, error) {
  if !s.valuesA2.isCached() {
    s.valuesA2.set("String")
  }
  return s.valuesA2.get()
}

func (s *spreadsheet) ValuesB2() (interface{}, error) {
  if !s.valuesB2.isCached() {
    s.valuesB2.set("String")
  }
  return s.valuesB2.get()
}

func (s *spreadsheet) ValuesA3() (interface{}, error) {
  if !s.valuesA3.isCached() {
    s.valuesA3.set("Integer")
  }
  return s.valuesA3.get()
}

func (s *spreadsheet) ValuesB3() (interface{}, error) {
  if !s.valuesB3.isCached() {
    s.valuesB3.set(1.0)
  }
  return s.valuesB3.get()
}


func (s *spreadsheet) SetvaluesB3(v interface{}) {
    s.valuesB3.set(v)
}

