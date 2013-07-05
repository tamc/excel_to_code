require_relative '../../spec_helper'

describe CompileToRubyUnitTest do
  
  def compile(text, sloppy = false)
    input = StringIO.new(text)
    output = StringIO.new
    CompileToRubyUnitTest.rewrite(input, sloppy, output)
    output.string
  end
  
it "should compile basic values and give precise tests when sloppy = false" do

input = <<END
sheet1\tA1\t[:number, "1"]
sheet1\tA2\t[:string, "Hello"]
sheet1\tA3\t[:error, "#NAME?"]
sheet1\tA4\t[:boolean_true]
sheet1\tA5\t[:boolean_false]
sheet1\tA6\t[:blank]
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
sheet1\tA1\t[:number, "1000"]
sheet1\tA2\t[:number, "0.1"]
sheet1\tA3\t[:number, "0"]
sheet1\tA4\t[:string, "Hello"]
sheet1\tA5\t[:error, "#NAME?"]
sheet1\tA6\t[:boolean_true]
sheet1\tA7\t[:boolean_false]
sheet1\tA8\t[:blank]
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
  lambda { compile("sheet1\tA1\t[:unknown]")}.should raise_exception(NotSupportedException)
end

end

