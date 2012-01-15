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
  def test_a1; assert_equal(1,worksheet.a1); end
  def test_a2; assert_equal("Hello",worksheet.a2); end
  def test_a3; assert_equal(:name,worksheet.a3); end
  def test_a4; assert_equal(true,worksheet.a4); end
  def test_a5; assert_equal(false,worksheet.a5); end
END

compile(input).should == expected
end

it "should raise an exception when values types are not recognised" do
  lambda { compile("A1\t[:unkown]")}.should raise_exception(NotSupportedException)
end

end

