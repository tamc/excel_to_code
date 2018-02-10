// Compiled version of /Users/tamc/Documents/github/excel_to_code/spec/test.xlsx
package excelspreadsheet

import "./excel"

type spreadsheet struct {
  valuesA1 excel.CachedValue
  valuesB1 excel.CachedValue
  valuesA2 excel.CachedValue
  valuesB2 excel.CachedValue
  valuesA3 excel.CachedValue
  valuesB3 excel.CachedValue
  valuesA4 excel.CachedValue
  valuesB4 excel.CachedValue
  valuesA5 excel.CachedValue
  valuesB5 excel.CachedValue
  valuesA6 excel.CachedValue
  valuesB6 excel.CachedValue
  valuesA7 excel.CachedValue
  valuesB7 excel.CachedValue
  valuesA8 excel.CachedValue
  valuesB8 excel.CachedValue
  valuesA9 excel.CachedValue
  valuesB9 excel.CachedValue
  valuesA10 excel.CachedValue
  valuesB10 excel.CachedValue
  valuesA11 excel.CachedValue
  valuesB11 excel.CachedValue
  valuesA12 excel.CachedValue
  valuesB12 excel.CachedValue
  valuesC1 excel.CachedValue
}

func New() spreadsheet {
  return spreadsheet{}
}

func (s *spreadsheet) ValuesA1() (interface{}, error) {
  if !s.valuesA1.IsCached() {
    s.valuesA1.Set("String")
  }
  return s.valuesA1.Get()
}

func (s *spreadsheet) ValuesB1() (interface{}, error) {
  if !s.valuesB1.IsCached() {
    s.valuesB1.Set("String")
  }
  return s.valuesB1.Get()
}

func (s *spreadsheet) ValuesA2() (interface{}, error) {
  if !s.valuesA2.IsCached() {
    s.valuesA2.Set("String")
  }
  return s.valuesA2.Get()
}

func (s *spreadsheet) ValuesB2() (interface{}, error) {
  if !s.valuesB2.IsCached() {
    s.valuesB2.Set("String")
  }
  return s.valuesB2.Get()
}

func (s *spreadsheet) ValuesA3() (interface{}, error) {
  if !s.valuesA3.IsCached() {
    s.valuesA3.Set("Integer")
  }
  return s.valuesA3.Get()
}

func (s *spreadsheet) ValuesB3() (interface{}, error) {
  if !s.valuesB3.IsCached() {
    s.valuesB3.Set(1.0)
  }
  return s.valuesB3.Get()
}

func (s *spreadsheet) ValuesA4() (interface{}, error) {
  if !s.valuesA4.IsCached() {
    s.valuesA4.Set("Float")
  }
  return s.valuesA4.Get()
}

func (s *spreadsheet) ValuesB4() (interface{}, error) {
  if !s.valuesB4.IsCached() {
    s.valuesB4.Set(1.1)
  }
  return s.valuesB4.Get()
}

func (s *spreadsheet) ValuesA5() (interface{}, error) {
  if !s.valuesA5.IsCached() {
    s.valuesA5.Set("Percent")
  }
  return s.valuesA5.Get()
}

func (s *spreadsheet) ValuesB5() (interface{}, error) {
  if !s.valuesB5.IsCached() {
    s.valuesB5.Set(0.999)
  }
  return s.valuesB5.Get()
}

func (s *spreadsheet) ValuesA6() (interface{}, error) {
  if !s.valuesA6.IsCached() {
    s.valuesA6.Set("Boolean")
  }
  return s.valuesA6.Get()
}

func (s *spreadsheet) ValuesB6() (interface{}, error) {
  if !s.valuesB6.IsCached() {
    s.valuesB6.Set(true)
  }
  return s.valuesB6.Get()
}

func (s *spreadsheet) ValuesA7() (interface{}, error) {
  if !s.valuesA7.IsCached() {
    s.valuesA7.Set("Boolean")
  }
  return s.valuesA7.Get()
}

func (s *spreadsheet) ValuesB7() (interface{}, error) {
  if !s.valuesB7.IsCached() {
    s.valuesB7.Set(false)
  }
  return s.valuesB7.Get()
}

func (s *spreadsheet) ValuesA8() (interface{}, error) {
  if !s.valuesA8.IsCached() {
    s.valuesA8.Set("Error")
  }
  return s.valuesA8.Get()
}

func (s *spreadsheet) ValuesB8() (interface{}, error) {
  if !s.valuesB8.IsCached() {
    s.valuesB8.Set(excel.Div0Error{})
  }
  return s.valuesB8.Get()
}

func (s *spreadsheet) ValuesA9() (interface{}, error) {
  if !s.valuesA9.IsCached() {
    s.valuesA9.Set("Error")
  }
  return s.valuesA9.Get()
}

func (s *spreadsheet) ValuesB9() (interface{}, error) {
  if !s.valuesB9.IsCached() {
    s.valuesB9.Set(0.0)
  }
  return s.valuesB9.Get()
}

func (s *spreadsheet) ValuesA10() (interface{}, error) {
  if !s.valuesA10.IsCached() {
    s.valuesA10.Set("Error")
  }
  return s.valuesA10.Get()
}

func (s *spreadsheet) ValuesB10() (interface{}, error) {
  if !s.valuesB10.IsCached() {
    s.valuesB10.Set(excel.NAError{})
  }
  return s.valuesB10.Get()
}

func (s *spreadsheet) ValuesA11() (interface{}, error) {
  if !s.valuesA11.IsCached() {
    s.valuesA11.Set("Error")
  }
  return s.valuesA11.Get()
}

func (s *spreadsheet) ValuesB11() (interface{}, error) {
  if !s.valuesB11.IsCached() {
    s.valuesB11.Set(excel.NameError{})
  }
  return s.valuesB11.Get()
}

func (s *spreadsheet) ValuesA12() (interface{}, error) {
  if !s.valuesA12.IsCached() {
    s.valuesA12.Set("Error")
  }
  return s.valuesA12.Get()
}

func (s *spreadsheet) ValuesB12() (interface{}, error) {
  if !s.valuesB12.IsCached() {
    s.valuesB12.Set(excel.NameError{})
  }
  return s.valuesB12.Get()
}

func (s *spreadsheet) ValuesC1() (interface{}, error) {
  if !s.valuesC1.IsCached() {
    s.valuesC1.Set(excel.Blank{})
  }
  return s.valuesC1.Get()
}


