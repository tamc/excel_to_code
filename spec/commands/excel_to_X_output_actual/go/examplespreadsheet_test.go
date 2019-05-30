// Test of compiled version of /Users/tamc/Documents/excel_to_code/spec/test_data/GoTestSpreadsheet.xlsx
package examplespreadsheet

import (
    "testing"
)

func TestValuetypesA1(t *testing.T) {
  s := New()
  e := "The basic types"
  a, err := s.ValuetypesA1()
  if a != e || err != nil {
      t.Errorf("ValuetypesA1 = (%v, %v), want (%v, nil)", a, err, e)
  }
}

func TestValuetypesB1(t *testing.T) {
  s := New()
  e := "The error types"
  a, err := s.ValuetypesB1()
  if a != e || err != nil {
      t.Errorf("ValuetypesB1 = (%v, %v), want (%v, nil)", a, err, e)
  }
}

func TestValuetypesA2(t *testing.T) {
  s := New()
  e := true
  a, err := s.ValuetypesA2()
  if a != e || err != nil {
      t.Errorf("ValuetypesA2 = (%v, %v), want (%v, nil)", a, err, e)
  }
}

func TestValuetypesB2(t *testing.T) {
  s := New()
  e := NAError{}
  a, err := s.ValuetypesB2()
  if err != e {
      t.Errorf("ValuetypesB2 = (%v, %v), want (nil, %v)", a, err, e)
  }
}
func TestValuetypesA3(t *testing.T) {
  s := New()
  e := false
  a, err := s.ValuetypesA3()
  if a != e || err != nil {
      t.Errorf("ValuetypesA3 = (%v, %v), want (%v, nil)", a, err, e)
  }
}

func TestValuetypesB3(t *testing.T) {
  s := New()
  e := NameError{}
  a, err := s.ValuetypesB3()
  if err != e {
      t.Errorf("ValuetypesB3 = (%v, %v), want (nil, %v)", a, err, e)
  }
}
func TestValuetypesA4(t *testing.T) {
  s := New()
  e := 123.0
  a, err := s.ValuetypesA4()
  if a != e || err != nil {
      t.Errorf("ValuetypesA4 = (%v, %v), want (%v, nil)", a, err, e)
  }
}

func TestValuetypesB4(t *testing.T) {
  s := New()
  e := Div0Error{}
  a, err := s.ValuetypesB4()
  if err != e {
      t.Errorf("ValuetypesB4 = (%v, %v), want (nil, %v)", a, err, e)
  }
}
func TestValuetypesA5(t *testing.T) {
  s := New()
  e := 3.1415
  a, err := s.ValuetypesA5()
  if a != e || err != nil {
      t.Errorf("ValuetypesA5 = (%v, %v), want (%v, nil)", a, err, e)
  }
}

func TestValuetypesB5(t *testing.T) {
  s := New()
  e := ValueError{}
  a, err := s.ValuetypesB5()
  if err != e {
      t.Errorf("ValuetypesB5 = (%v, %v), want (nil, %v)", a, err, e)
  }
}
func TestValuetypesA6(t *testing.T) {
  s := New()
  e := "Hello world"
  a, err := s.ValuetypesA6()
  if a != e || err != nil {
      t.Errorf("ValuetypesA6 = (%v, %v), want (%v, nil)", a, err, e)
  }
}

func TestValuetypesB6(t *testing.T) {
  s := New()
  e := Infinity
  a, err := s.ValuetypesB6()
  if a != e || err != nil {
      t.Errorf("ValuetypesB6 = (%v, %v), want (%v, nil)", a, err, e)
  }
}


