// Compiled version of /Users/tamc/Documents/excel_to_code/spec/test_data/GoTestSpreadsheet.xlsx
package examplespreadsheet

import (
	"fmt"
	"strconv"
)

type spreadsheet struct {
valuetypesA1 cachedValue
valuetypesB1 cachedValue
valuetypesA2 cachedValue
valuetypesB2 cachedValue
valuetypesA3 cachedValue
valuetypesB3 cachedValue
valuetypesA4 cachedValue
valuetypesB4 cachedValue
valuetypesA5 cachedValue
valuetypesB5 cachedValue
valuetypesA6 cachedValue
valuetypesB6 cachedValue
}
 func New() spreadsheet {
  return spreadsheet{}
}
func (s *spreadsheet) ValuetypesA1() (interface{}, error) {
  if !s.valuetypesA1.isCached() {
    s.valuetypesA1.set("The basic types")
  }
  return s.valuetypesA1.get()
}
func (s *spreadsheet) ValuetypesB1() (interface{}, error) {
  if !s.valuetypesB1.isCached() {
    s.valuetypesB1.set("The error types")
  }
  return s.valuetypesB1.get()
}
func (s *spreadsheet) ValuetypesA2() (interface{}, error) {
  if !s.valuetypesA2.isCached() {
    s.valuetypesA2.set(true)
  }
  return s.valuetypesA2.get()
}
func (s *spreadsheet) ValuetypesB2() (interface{}, error) {
  if !s.valuetypesB2.isCached() {
    s.valuetypesB2.set(NAError{})
  }
  return s.valuetypesB2.get()
}
func (s *spreadsheet) ValuetypesA3() (interface{}, error) {
  if !s.valuetypesA3.isCached() {
    s.valuetypesA3.set(false)
  }
  return s.valuetypesA3.get()
}
func (s *spreadsheet) ValuetypesB3() (interface{}, error) {
  if !s.valuetypesB3.isCached() {
    s.valuetypesB3.set(NameError{})
  }
  return s.valuetypesB3.get()
}
func (s *spreadsheet) ValuetypesA4() (interface{}, error) {
  if !s.valuetypesA4.isCached() {
    s.valuetypesA4.set(123.0)
  }
  return s.valuetypesA4.get()
}
func (s *spreadsheet) ValuetypesB4() (interface{}, error) {
  if !s.valuetypesB4.isCached() {
    s.valuetypesB4.set(Div0Error{})
  }
  return s.valuetypesB4.get()
}
func (s *spreadsheet) ValuetypesA5() (interface{}, error) {
  if !s.valuetypesA5.isCached() {
    s.valuetypesA5.set(3.1415)
  }
  return s.valuetypesA5.get()
}
func (s *spreadsheet) ValuetypesB5() (interface{}, error) {
  if !s.valuetypesB5.isCached() {
    s.valuetypesB5.set(ValueError{})
  }
  return s.valuetypesB5.get()
}
func (s *spreadsheet) ValuetypesA6() (interface{}, error) {
  if !s.valuetypesA6.isCached() {
    s.valuetypesA6.set("Hello world")
  }
  return s.valuetypesA6.get()
}
func (s *spreadsheet) ValuetypesB6() (interface{}, error) {
  if !s.valuetypesB6.isCached() {
    s.valuetypesB6.set(NumError{})
  }
  return s.valuetypesB6.get()
}



// cachedValue wraps the different possible Excel Values
//
// Later, we may have one for each Excel type
// which is why we use wrappers.
type cachedValue struct {
	// v is the cached value, in this case of any type
	v interface{}
	// c is whether it is cached
	c bool
}

func (c *cachedValue) isCached() bool {
	return c.c
}

func (c *cachedValue) set(v interface{}) {
	c.v = v
	c.c = true
}

// get returns the cached value. If the cached
// value is an Excel Error that is returned as
// the second argument. If the item isn't cached
// then will return a NotCachedError.
func (c *cachedValue) get() (interface{}, error) {
	if !c.c {
		return nil, notCachedError{}
	}
	if err, ok := c.v.(error); ok {
		return nil, err
	}
	return c.v, nil
}

type notCachedError struct{}

func (e notCachedError) Error() string {
	return "Get() called but nothing cached"
}

// Blank is an empty Excel cell.
type Blank struct{}

// The Excel #VALUE! error
type ValueError struct {
	value interface{}
	cast  string
}

// The Excel #NAME? error
type NameError struct{}

// The Excel #DIV/0! error
type Div0Error struct{}

// The Excel #REF! error
type RefError struct{}

// The Excel #N/A error
type NAError struct{}

// The Excel #NUM! error
type NumError struct{}

func (e ValueError) Error() string {
	return fmt.Sprintf("#VALUE!: could not convert %v to %v", e.value, e.cast)
}

func (e NameError) Error() string {
	return "#NAME?"
}

func (e Div0Error) Error() string {
	return "#DIV/0!"
}

func (e RefError) Error() string {
	return "#REF!"
}

func (e NAError) Error() string {
	return "#N/A"
}

func (e NumError) Error() string {
	return "#DIV/0!"
}

func excel_if(c interface{}, t, f func() interface{}) interface{} {
	b, err := boolean(c)
	if err != nil {
		return err
	}
	if b {
		if t == nil {
			return true
		} else {
			return t()
		}
	} else {
		if f == nil {
			return false
		} else {
			return f()
		}
	}
}

func add(a, b interface{}) interface{} {
	n1, err := number(a)
	if err != nil {
		return err
	}
	n2, err := number(b)
	if err != nil {
		return err
	}
	return n1 + n2
}

func number(a interface{}) (float64, error) {
	switch a := a.(type) {
	case Blank:
		return 0, nil
	case float64:
		return a, nil
	case bool:
		if a {
			return 1, nil
		} else {
			return 0, nil
		}
	case error:
		return 0, a
	case string:
		i, err := strconv.ParseFloat(a, 64)
		if err != nil {
			return 0, ValueError{a, "float64"}
		}
		return i, nil
	default:
		return 0, ValueError{a, "float64"}
	}
}

func boolean(a interface{}) (bool, error) {
	switch a := a.(type) {
	case Blank:
		return false, nil
	case float64:
		if a == 0 {
			return false, nil
		} else {
			return true, nil
		}
	case bool:
		return a, nil
	case error:
		return false, a
	default:
		return false, ValueError{a, "bool"}
	}
}

