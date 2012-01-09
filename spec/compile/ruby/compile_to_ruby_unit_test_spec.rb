require_relative '../../spec_helper'

describe CompileToRubyUnitTest do
  
  def compile(text)
    input = StringIO.new(text)
    output = StringIO.new
    CompileToRubyUnitTest.rewrite(input,output)
    output.string
  end
  
it "should compile basic values" do
input = <<END
A1\t[:number, "1"]
A2\t[:string, "Hello"]
A3\t[:error, "#NAME?"]
A4\t[:boolean_true]
A5\t[:boolean_false]
END

expected = <<END
  def test_a1; assert_equal(worksheet.a1,1); end
  def test_a2; assert_equal(worksheet.a2,"Hello"); end
  def test_a3; assert_equal(worksheet.a3,:name); end
  def test_a4; assert_equal(worksheet.a4,true); end
  def test_a5; assert_equal(worksheet.a5,false); end
END

compile(input).should == expected
end

it "should raise an exception when values types are not recognised" do
  lambda { compile("A1\t[:unkown]")}.should raise_exception(NotSupportedException)
end

end

