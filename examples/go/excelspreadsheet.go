// Compiled version of /Users/tamc/Documents/github/excel_to_code/spec/test.xlsx
package excelspreadsheet

import "./excel"

type spreadsheet struct {
  valuesA1 interface{}
  valuesB1 interface{}
  valuesA2 interface{}
  valuesB2 interface{}
  valuesA3 interface{}
  valuesB3 interface{}
  valuesA4 interface{}
  valuesB4 interface{}
  valuesA5 interface{}
  valuesB5 interface{}
  valuesA6 interface{}
  valuesB6 interface{}
  valuesA7 interface{}
  valuesB7 interface{}
  valuesA8 interface{}
  valuesB8 interface{}
  valuesA9 interface{}
  valuesB9 interface{}
  valuesA10 interface{}
  valuesB10 interface{}
  valuesA11 interface{}
  valuesB11 interface{}
  valuesA12 interface{}
  valuesB12 interface{}
  valuesC1 interface{}
}

func New() spreadsheet {
  return spreadsheet{}
}

func (s *spreadsheet) ValuesA1() interface{} {
  if s.valuesA1 == nil {
    s.valuesA1 = "String"
  }
  return s.valuesA1
}

func (s *spreadsheet) ValuesB1() interface{} {
  if s.valuesB1 == nil {
    s.valuesB1 = "String"
  }
  return s.valuesB1
}

func (s *spreadsheet) ValuesA2() interface{} {
  if s.valuesA2 == nil {
    s.valuesA2 = "String"
  }
  return s.valuesA2
}

func (s *spreadsheet) ValuesB2() interface{} {
  if s.valuesB2 == nil {
    s.valuesB2 = "String"
  }
  return s.valuesB2
}

func (s *spreadsheet) ValuesA3() interface{} {
  if s.valuesA3 == nil {
    s.valuesA3 = "Integer"
  }
  return s.valuesA3
}

func (s *spreadsheet) ValuesB3() interface{} {
  if s.valuesB3 == nil {
    s.valuesB3 = 1.0
  }
  return s.valuesB3
}

func (s *spreadsheet) ValuesA4() interface{} {
  if s.valuesA4 == nil {
    s.valuesA4 = "Float"
  }
  return s.valuesA4
}

func (s *spreadsheet) ValuesB4() interface{} {
  if s.valuesB4 == nil {
    s.valuesB4 = 1.1
  }
  return s.valuesB4
}

func (s *spreadsheet) ValuesA5() interface{} {
  if s.valuesA5 == nil {
    s.valuesA5 = "Percent"
  }
  return s.valuesA5
}

func (s *spreadsheet) ValuesB5() interface{} {
  if s.valuesB5 == nil {
    s.valuesB5 = 0.999
  }
  return s.valuesB5
}

func (s *spreadsheet) ValuesA6() interface{} {
  if s.valuesA6 == nil {
    s.valuesA6 = "Boolean"
  }
  return s.valuesA6
}

func (s *spreadsheet) ValuesB6() interface{} {
  if s.valuesB6 == nil {
    s.valuesB6 = true
  }
  return s.valuesB6
}

func (s *spreadsheet) ValuesA7() interface{} {
  if s.valuesA7 == nil {
    s.valuesA7 = "Boolean"
  }
  return s.valuesA7
}

func (s *spreadsheet) ValuesB7() interface{} {
  if s.valuesB7 == nil {
    s.valuesB7 = false
  }
  return s.valuesB7
}

func (s *spreadsheet) ValuesA8() interface{} {
  if s.valuesA8 == nil {
    s.valuesA8 = "Error"
  }
  return s.valuesA8
}

func (s *spreadsheet) ValuesB8() interface{} {
  if s.valuesB8 == nil {
    s.valuesB8 = excel.Div0Error{}
  }
  return s.valuesB8
}

func (s *spreadsheet) ValuesA9() interface{} {
  if s.valuesA9 == nil {
    s.valuesA9 = "Error"
  }
  return s.valuesA9
}

func (s *spreadsheet) ValuesB9() interface{} {
  if s.valuesB9 == nil {
    s.valuesB9 = 0.0
  }
  return s.valuesB9
}

func (s *spreadsheet) ValuesA10() interface{} {
  if s.valuesA10 == nil {
    s.valuesA10 = "Error"
  }
  return s.valuesA10
}

func (s *spreadsheet) ValuesB10() interface{} {
  if s.valuesB10 == nil {
    s.valuesB10 = excel.NAError{}
  }
  return s.valuesB10
}

func (s *spreadsheet) ValuesA11() interface{} {
  if s.valuesA11 == nil {
    s.valuesA11 = "Error"
  }
  return s.valuesA11
}

func (s *spreadsheet) ValuesB11() interface{} {
  if s.valuesB11 == nil {
    s.valuesB11 = excel.NameError{}
  }
  return s.valuesB11
}

func (s *spreadsheet) ValuesA12() interface{} {
  if s.valuesA12 == nil {
    s.valuesA12 = "Error"
  }
  return s.valuesA12
}

func (s *spreadsheet) ValuesB12() interface{} {
  if s.valuesB12 == nil {
    s.valuesB12 = excel.NameError{}
  }
  return s.valuesB12
}

func (s *spreadsheet) ValuesC1() interface{} {
  if s.valuesC1 == nil {
    s.valuesC1 = excel.Blank{}
  }
  return s.valuesC1
}


