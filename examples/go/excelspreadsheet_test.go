// Test of compiled version of /Users/tamc/Documents/github/excel_to_code/spec/test.xlsx
package excelspreadsheet

import (
    "testing"
)

func TestValuesA1(t *testing.T) {
  s := New()
  e := "String"
  a, err := s.ValuesA1()
  if a != e || err != nil {
      t.Errorf("ValuesA1 = (%v, %v), want (%v, nil)", a, err, e)
  }
}

func TestValuesB1(t *testing.T) {
  s := New()
  e := "String"
  a, err := s.ValuesB1()
  if a != e || err != nil {
      t.Errorf("ValuesB1 = (%v, %v), want (%v, nil)", a, err, e)
  }
}

func TestValuesA2(t *testing.T) {
  s := New()
  e := "String"
  a, err := s.ValuesA2()
  if a != e || err != nil {
      t.Errorf("ValuesA2 = (%v, %v), want (%v, nil)", a, err, e)
  }
}

func TestValuesB2(t *testing.T) {
  s := New()
  e := "String"
  a, err := s.ValuesB2()
  if a != e || err != nil {
      t.Errorf("ValuesB2 = (%v, %v), want (%v, nil)", a, err, e)
  }
}

func TestValuesA3(t *testing.T) {
  s := New()
  e := "Integer"
  a, err := s.ValuesA3()
  if a != e || err != nil {
      t.Errorf("ValuesA3 = (%v, %v), want (%v, nil)", a, err, e)
  }
}

func TestValuesB3(t *testing.T) {
  s := New()
  e := 1.0
  a, err := s.ValuesB3()
  if a != e || err != nil {
      t.Errorf("ValuesB3 = (%v, %v), want (%v, nil)", a, err, e)
  }
}

func TestValuesA4(t *testing.T) {
  s := New()
  e := "Float"
  a, err := s.ValuesA4()
  if a != e || err != nil {
      t.Errorf("ValuesA4 = (%v, %v), want (%v, nil)", a, err, e)
  }
}

func TestValuesB4(t *testing.T) {
  s := New()
  e := 1.1
  a, err := s.ValuesB4()
  if a != e || err != nil {
      t.Errorf("ValuesB4 = (%v, %v), want (%v, nil)", a, err, e)
  }
}

func TestValuesA5(t *testing.T) {
  s := New()
  e := "Percent"
  a, err := s.ValuesA5()
  if a != e || err != nil {
      t.Errorf("ValuesA5 = (%v, %v), want (%v, nil)", a, err, e)
  }
}

func TestValuesB5(t *testing.T) {
  s := New()
  e := 0.999
  a, err := s.ValuesB5()
  if a != e || err != nil {
      t.Errorf("ValuesB5 = (%v, %v), want (%v, nil)", a, err, e)
  }
}

func TestValuesA6(t *testing.T) {
  s := New()
  e := "Boolean"
  a, err := s.ValuesA6()
  if a != e || err != nil {
      t.Errorf("ValuesA6 = (%v, %v), want (%v, nil)", a, err, e)
  }
}

func TestValuesB6(t *testing.T) {
  s := New()
  e := true
  a, err := s.ValuesB6()
  if a != e || err != nil {
      t.Errorf("ValuesB6 = (%v, %v), want (%v, nil)", a, err, e)
  }
}

func TestValuesA7(t *testing.T) {
  s := New()
  e := "Boolean"
  a, err := s.ValuesA7()
  if a != e || err != nil {
      t.Errorf("ValuesA7 = (%v, %v), want (%v, nil)", a, err, e)
  }
}

func TestValuesB7(t *testing.T) {
  s := New()
  e := false
  a, err := s.ValuesB7()
  if a != e || err != nil {
      t.Errorf("ValuesB7 = (%v, %v), want (%v, nil)", a, err, e)
  }
}

func TestValuesA8(t *testing.T) {
  s := New()
  e := "Error"
  a, err := s.ValuesA8()
  if a != e || err != nil {
      t.Errorf("ValuesA8 = (%v, %v), want (%v, nil)", a, err, e)
  }
}

func TestValuesB8(t *testing.T) {
  s := New()
  e := Div0Error{}
  a, err := s.ValuesB8()
  if err != e {
      t.Errorf("ValuesB8 = (%v, %v), want (nil, %v)", a, err, e)
  }
}

func TestValuesA9(t *testing.T) {
  s := New()
  e := "Error"
  a, err := s.ValuesA9()
  if a != e || err != nil {
      t.Errorf("ValuesA9 = (%v, %v), want (%v, nil)", a, err, e)
  }
}

func TestValuesB9(t *testing.T) {
  s := New()
  e := 0.0
  a, err := s.ValuesB9()
  if a != e || err != nil {
      t.Errorf("ValuesB9 = (%v, %v), want (%v, nil)", a, err, e)
  }
}

func TestValuesA10(t *testing.T) {
  s := New()
  e := "Error"
  a, err := s.ValuesA10()
  if a != e || err != nil {
      t.Errorf("ValuesA10 = (%v, %v), want (%v, nil)", a, err, e)
  }
}

func TestValuesB10(t *testing.T) {
  s := New()
  e := NAError{}
  a, err := s.ValuesB10()
  if err != e {
      t.Errorf("ValuesB10 = (%v, %v), want (nil, %v)", a, err, e)
  }
}

func TestValuesA11(t *testing.T) {
  s := New()
  e := "Error"
  a, err := s.ValuesA11()
  if a != e || err != nil {
      t.Errorf("ValuesA11 = (%v, %v), want (%v, nil)", a, err, e)
  }
}

func TestValuesB11(t *testing.T) {
  s := New()
  e := NameError{}
  a, err := s.ValuesB11()
  if err != e {
      t.Errorf("ValuesB11 = (%v, %v), want (nil, %v)", a, err, e)
  }
}

func TestValuesA12(t *testing.T) {
  s := New()
  e := "Error"
  a, err := s.ValuesA12()
  if a != e || err != nil {
      t.Errorf("ValuesA12 = (%v, %v), want (%v, nil)", a, err, e)
  }
}

func TestValuesB12(t *testing.T) {
  s := New()
  e := NameError{}
  a, err := s.ValuesB12()
  if err != e {
      t.Errorf("ValuesB12 = (%v, %v), want (nil, %v)", a, err, e)
  }
}

func TestValuesC1(t *testing.T) {
  s := New()
  e := Blank{}
  a, err := s.ValuesC1()
  if a != e || err != nil {
      t.Errorf("ValuesC1 = (%v, %v), want (%v, nil)", a, err, e)
  }
}


