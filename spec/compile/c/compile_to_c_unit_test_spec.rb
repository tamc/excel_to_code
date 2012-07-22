require_relative '../../spec_helper'

describe CompileToCUnitTest do
  
  def compile(text, sloppy = false)
    input = StringIO.new(text)
    output = StringIO.new
    CompileToCUnitTest.rewrite(input, sloppy, 'sheet1', ['A1','A2','A3','A4','A5','A6','A7','A8'] , output)
    output.string
  end
  
it "should compile basic values and give precise tests when sloppy = false" do

input = <<END
A1\t[:number, "1"]
A2\t[:string, "Hello"]
A3\t[:error, "#NAME?"]
A4\t[:boolean_true]
A5\t[:boolean_false]
A6\t[:blank]
END

expected = <<END
def test_sheet1_a1
  r = spreadsheet.sheet1_a1
  assert_equal(:ExcelNumber,r[:type])
  assert_equal(1.0,r[:number])
end

def test_sheet1_a2
  r = spreadsheet.sheet1_a2
  assert_equal(:ExcelString,r[:type])
  assert_equal("Hello",r[:string].force_encoding('utf-8'))
end

def test_sheet1_a3
  r = spreadsheet.sheet1_a3
  assert_equal(:ExcelError,r[:type])
end

def test_sheet1_a4
  r = spreadsheet.sheet1_a4
  assert_equal(:ExcelBoolean,r[:type])
  assert_equal(1,r[:number])
end

def test_sheet1_a5
  r = spreadsheet.sheet1_a5
  assert_equal(:ExcelBoolean,r[:type])
  assert_equal(0,r[:number])
end

def test_sheet1_a6
  r = spreadsheet.sheet1_a6
  assert_equal(:ExcelEmpty,r[:type])
end

END
compile(input).should == expected
end

it "should compile basic values and give less precise tests when sloppy = true" do

input = <<END
A1\t[:number, "1000"]
A2\t[:number, "0.1"]
A3\t[:number, "0"]
A4\t[:string, "Hello"]
A5\t[:error, "#NAME?"]
A6\t[:boolean_true]
A7\t[:boolean_false]
A8\t[:blank]
END

expected = <<END
def test_sheet1_a1
  r = spreadsheet.sheet1_a1
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(1000.0,r[:number],0.001)
end

def test_sheet1_a2
  r = spreadsheet.sheet1_a2
  assert_equal(:ExcelNumber,r[:type])
  assert_in_delta(0.1,r[:number],0.001)
end

def test_sheet1_a3
  r = spreadsheet.sheet1_a3
  pass if r[:type] == :ExcelEmpty
  assert_equal(:ExcelNumber,r[:type])
  assert_in_delta(0.0,r[:number],0.001)
end

def test_sheet1_a4
  r = spreadsheet.sheet1_a4
  assert_equal(:ExcelString,r[:type])
  assert_equal("Hello",r[:string].force_encoding('utf-8'))
end

def test_sheet1_a5
  r = spreadsheet.sheet1_a5
  assert_equal(:ExcelError,r[:type])
end

def test_sheet1_a6
  r = spreadsheet.sheet1_a6
  assert_equal(:ExcelBoolean,r[:type])
  assert_equal(1,r[:number])
end

def test_sheet1_a7
  r = spreadsheet.sheet1_a7
  assert_equal(:ExcelBoolean,r[:type])
  assert_equal(0,r[:number])
end

def test_sheet1_a8
  r = spreadsheet.sheet1_a8
  pass if r[:type] == :ExcelEmpty
  assert_equal(:ExcelNumber,r[:type])
  assert_in_delta(0.0,r[:number],0.001)
end

END

compile(input, true).should == expected
end

it "should raise an exception when values types are not recognised" do
  lambda { compile("A1\t[:unknown]")}.should raise_exception(NotSupportedException)
end

end

