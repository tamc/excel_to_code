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
  def test_sheet1_a1; assert_equal(1, worksheet.sheet1_a1); end
  def test_sheet1_a2; assert_equal("Hello", worksheet.sheet1_a2); end
  def test_sheet1_a3; assert_equal(:name, worksheet.sheet1_a3); end
  def test_sheet1_a4; assert_equal(true, worksheet.sheet1_a4); end
  def test_sheet1_a5; assert_equal(false, worksheet.sheet1_a5); end
  def test_sheet1_a6; assert_equal(nil, worksheet.sheet1_a6); end
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
  def test_sheet1_a1; assert_in_epsilon(1000, worksheet.sheet1_a1, 0.001); end
  def test_sheet1_a2; assert_in_delta(0.1, worksheet.sheet1_a2, 0.001); end
  def test_sheet1_a3; assert_in_delta(0, (worksheet.sheet1_a3||0), 0.001); end
  def test_sheet1_a4; assert_equal("Hello", worksheet.sheet1_a4); end
  def test_sheet1_a5; assert_equal(:name, worksheet.sheet1_a5); end
  def test_sheet1_a6; assert_equal(true, worksheet.sheet1_a6); end
  def test_sheet1_a7; assert_equal(false, worksheet.sheet1_a7); end
  def test_sheet1_a8; assert_includes([nil, 0], worksheet.sheet1_a8); end
END

compile(input, true).should == expected
end

it "should raise an exception when values types are not recognised" do
  lambda { compile("A1\t[:unknown]")}.should raise_exception(NotSupportedException)
end

end

