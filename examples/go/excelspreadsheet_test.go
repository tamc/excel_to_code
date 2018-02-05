// Test of compiled version of /Users/tamc/Documents/github/excel_to_code/spec/test.xlsx
package excelspreadsheet

import (
    "./excel"
    "testing"
)

func TestValuesA1(t *testing.T) {
  s := New()
  e := "String"
  a := s.ValuesA1()
  if a != e {
      t.Errorf("ValuesA1 = %v, want %v", a, e)
  }
 }

func TestValuesB1(t *testing.T) {
  s := New()
  e := "String"
  a := s.ValuesB1()
  if a != e {
      t.Errorf("ValuesB1 = %v, want %v", a, e)
  }
 }

func TestValuesA2(t *testing.T) {
  s := New()
  e := "String"
  a := s.ValuesA2()
  if a != e {
      t.Errorf("ValuesA2 = %v, want %v", a, e)
  }
 }

func TestValuesB2(t *testing.T) {
  s := New()
  e := "String"
  a := s.ValuesB2()
  if a != e {
      t.Errorf("ValuesB2 = %v, want %v", a, e)
  }
 }

func TestValuesA3(t *testing.T) {
  s := New()
  e := "Integer"
  a := s.ValuesA3()
  if a != e {
      t.Errorf("ValuesA3 = %v, want %v", a, e)
  }
 }

func TestValuesB3(t *testing.T) {
  s := New()
  e := 1.0
  a := s.ValuesB3()
  if a != e {
      t.Errorf("ValuesB3 = %v, want %v", a, e)
  }
 }

func TestValuesA4(t *testing.T) {
  s := New()
  e := "Float"
  a := s.ValuesA4()
  if a != e {
      t.Errorf("ValuesA4 = %v, want %v", a, e)
  }
 }

func TestValuesB4(t *testing.T) {
  s := New()
  e := 1.1
  a := s.ValuesB4()
  if a != e {
      t.Errorf("ValuesB4 = %v, want %v", a, e)
  }
 }

func TestValuesA5(t *testing.T) {
  s := New()
  e := "Percent"
  a := s.ValuesA5()
  if a != e {
      t.Errorf("ValuesA5 = %v, want %v", a, e)
  }
 }

func TestValuesB5(t *testing.T) {
  s := New()
  e := 0.999
  a := s.ValuesB5()
  if a != e {
      t.Errorf("ValuesB5 = %v, want %v", a, e)
  }
 }

func TestValuesA6(t *testing.T) {
  s := New()
  e := "Boolean"
  a := s.ValuesA6()
  if a != e {
      t.Errorf("ValuesA6 = %v, want %v", a, e)
  }
 }

func TestValuesB6(t *testing.T) {
  s := New()
  e := true
  a := s.ValuesB6()
  if a != e {
      t.Errorf("ValuesB6 = %v, want %v", a, e)
  }
 }

func TestValuesA7(t *testing.T) {
  s := New()
  e := "Boolean"
  a := s.ValuesA7()
  if a != e {
      t.Errorf("ValuesA7 = %v, want %v", a, e)
  }
 }

func TestValuesB7(t *testing.T) {
  s := New()
  e := false
  a := s.ValuesB7()
  if a != e {
      t.Errorf("ValuesB7 = %v, want %v", a, e)
  }
 }

func TestValuesA8(t *testing.T) {
  s := New()
  e := "Error"
  a := s.ValuesA8()
  if a != e {
      t.Errorf("ValuesA8 = %v, want %v", a, e)
  }
 }

func TestValuesB8(t *testing.T) {
  s := New()
  e := excel.Div0Error{}
  a := s.ValuesB8()
  if a != e {
      t.Errorf("ValuesB8 = %v, want %v", a, e)
  }
 }

func TestValuesA9(t *testing.T) {
  s := New()
  e := "Error"
  a := s.ValuesA9()
  if a != e {
      t.Errorf("ValuesA9 = %v, want %v", a, e)
  }
 }

func TestValuesB9(t *testing.T) {
  s := New()
  e := 0.0
  a := s.ValuesB9()
  if a != e {
      t.Errorf("ValuesB9 = %v, want %v", a, e)
  }
 }

func TestValuesA10(t *testing.T) {
  s := New()
  e := "Error"
  a := s.ValuesA10()
  if a != e {
      t.Errorf("ValuesA10 = %v, want %v", a, e)
  }
 }

func TestValuesB10(t *testing.T) {
  s := New()
  e := excel.NAError{}
  a := s.ValuesB10()
  if a != e {
      t.Errorf("ValuesB10 = %v, want %v", a, e)
  }
 }

func TestValuesA11(t *testing.T) {
  s := New()
  e := "Error"
  a := s.ValuesA11()
  if a != e {
      t.Errorf("ValuesA11 = %v, want %v", a, e)
  }
 }

func TestValuesB11(t *testing.T) {
  s := New()
  e := excel.NameError{}
  a := s.ValuesB11()
  if a != e {
      t.Errorf("ValuesB11 = %v, want %v", a, e)
  }
 }

func TestValuesA12(t *testing.T) {
  s := New()
  e := "Error"
  a := s.ValuesA12()
  if a != e {
      t.Errorf("ValuesA12 = %v, want %v", a, e)
  }
 }

func TestValuesB12(t *testing.T) {
  s := New()
  e := excel.NameError{}
  a := s.ValuesB12()
  if a != e {
      t.Errorf("ValuesB12 = %v, want %v", a, e)
  }
 }

func TestValuesC1(t *testing.T) {
  s := New()
  e := excel.Blank{}
  a := s.ValuesC1()
  if a != e {
      t.Errorf("ValuesC1 = %v, want %v", a, e)
  }
 }


