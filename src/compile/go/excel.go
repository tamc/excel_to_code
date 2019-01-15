// excel replicates values and functions from Microsoft Excel
package excel

import (
	"fmt"
	"strconv"
)

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
