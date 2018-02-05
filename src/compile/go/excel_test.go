package excel

import "testing"

type Spreadsheet struct {
	inputA1  interface{}
	inputA2  interface{}
	outputA1 interface{}
}

func (s *Spreadsheet) InputA1() interface{} {
	if s.inputA1 == nil {
		s.inputA1 = Blank{}
	}
	return s.inputA1
}

func (s *Spreadsheet) SetInputA1(value interface{}) {
	s.inputA1 = value
}

func (s *Spreadsheet) InputA2() interface{} {
	if s.inputA2 == nil {
		s.inputA2 = Blank{}
	}
	return s.inputA2
}

func (s *Spreadsheet) SetInputA2(value interface{}) {
	s.inputA2 = value
}

func (s *Spreadsheet) OutputA1() interface{} {
	if s.outputA1 == nil {
		inputA1 := s.InputA1()
		inputA2 := s.InputA2()
		sum := add(inputA1, inputA2)
		s.outputA1 = excel_if(sum, func() interface{} {
			return Blank{}
		}, s.InputA1)
	}
	return s.outputA1
}

func TestHelloWorld(t *testing.T) {
	s := Spreadsheet{}
	s.SetInputA1(0.0)
	//inputA1 := s.InputA1()
	outputA1 := s.OutputA1()
	if outputA1 != 1 {
		t.Errorf("%v = %v, want %v", s.OutputA1, outputA1, 1)
	}
}
