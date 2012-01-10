require_relative '../../spec_helper'

describe CompileToRuby do
  
  def compile(text)
    input = StringIO.new(text)
    output = StringIO.new
    CompileToRuby.rewrite(input,output)
    output.string
  end
  
it "should compile simple arithmetic" do
input = <<END
A1\t[:arithmetic, [:number, "1"], [:operator, "+"], [:number, "1"]]
A2\t[:arithmetic, [:number, "1"], [:operator, "-"], [:number, "1"]]
A3\t[:arithmetic, [:number, "1"], [:operator, "*"], [:number, "1"]]
A4\t[:arithmetic, [:number, "1"], [:operator, "/"], [:number, "1"]]
A5\t[:arithmetic, [:number, "1"], [:operator, "^"], [:number, "1"]]
A6\t[:arithmetic, [:number, "1.1"], [:operator, "+"], [:number, "-1E12"]]
END

expected = <<END
  def a1; add(1,1); end
  def a2; subtract(1,1); end
  def a3; multiply(1,1); end
  def a4; divide(1,1); end
  def a5; power(1,1); end
  def a6; add(1.1,-1000000000000.0); end
END

compile(input).should == expected
end
end



