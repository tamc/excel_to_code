require_relative '../../spec_helper'

describe CompileToRubyUnitTest do
  
  def compile(input, sloppy = false, sheet_names = {})
    output = StringIO.new
    CompileToRubyUnitTest.rewrite(input, sloppy, sheet_names,  output)
    output.string
  end
  
it "should compile basic values and give precise tests when sloppy = false" do

  input = {
    ["sheet1", "A1"] => [:number, "1"],
    ["sheet1", "A2"] => [:string, "Hello"],
    ["sheet1", "A3"] => [:error, :"#NAME?"],
    ["sheet1", "A4"] => [:boolean_true],
    ["sheet1", "A5"] => [:boolean_false],
    ["sheet1", "A6"] => [:blank]
  }

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

input = {
  ["sheet1", "A1"] => [:number, "1000"],
  ["sheet1", "A2"] => [:number, "0.1"],
  ["sheet1", "A3"] => [:number, "0"],
  ["sheet1", "A4"] => [:string, "Hello"],
  ["sheet1", "A5"] => [:error, :"#NAME?"],
  ["sheet1", "A6"] => [:boolean_true],
  ["sheet1", "A7"] => [:boolean_false],
  ["sheet1", "A8"] => [:blank]
}

expected = <<END
  def test_sheet1_a1; assert_in_epsilon(1000, worksheet.sheet1_a1, 0.002); end
  def test_sheet1_a2; assert_in_delta(0.1, worksheet.sheet1_a2, 0.002); end
  def test_sheet1_a3; assert_in_delta(0, (worksheet.sheet1_a3||0), 0.002); end
  def test_sheet1_a4; assert_equal("Hello", worksheet.sheet1_a4); end
  def test_sheet1_a5; assert_equal(:name, worksheet.sheet1_a5); end
  def test_sheet1_a6; assert_equal(true, worksheet.sheet1_a6); end
  def test_sheet1_a7; assert_equal(false, worksheet.sheet1_a7); end
  def test_sheet1_a8; assert_includes([nil, 0], worksheet.sheet1_a8); end
END

compile(input, true).should == expected
end

it "should raise an exception when values types are not recognised" do
  lambda { compile({["sheet1","A1"] => [:unknown]})}.should raise_exception(NotSupportedException)
end

it "should map sheet names" do
input = {
  ["sheet with a silly name", "A1"] => [:number, "1000"],
}

sheet_names = { "sheet with a silly name" => 'sheet1' }

expected = <<END
  def test_sheet1_a1; assert_in_epsilon(1000, worksheet.sheet1_a1, 0.002); end
END

compile(input, true, sheet_names).should == expected
end

end

