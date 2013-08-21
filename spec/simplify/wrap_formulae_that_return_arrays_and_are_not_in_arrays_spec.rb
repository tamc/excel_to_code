require_relative '../spec_helper'

describe WrapFormulaeThatReturnArraysAndAReNotInArrays do
  
  it "should wrap formulae (such as MMULT(A1:B1,C1:C2)) that return arrays but are not inside an array formula" do

input = <<END
A1\t[:function, "MMULT", [:array, [:row, [:number, "1"]]], [:array, [:row, [:number, "2"]]]]
A2\t[:function, "INDEX", [:function, "MMULT", [:array, [:row, [:number, "1"]]], [:array, [:row, [:number, "2"]]]], [:number, "1"], [:number, "1"]]
END

expected_output = <<END
A1\t[:function, "INDEX", [:function, "MMULT", [:array, [:row, [:number, "1"]]], [:array, [:row, [:number, "2"]]]], [:number, "1"], [:number, "1"]]
A2\t[:function, "INDEX", [:function, "MMULT", [:array, [:row, [:number, "1"]]], [:array, [:row, [:number, "2"]]]], [:number, "1"], [:number, "1"]]
END
    
    input = StringIO.new(input)
    output = StringIO.new
    WrapFormulaeThatReturnArraysAndAReNotInArrays.replace(input,output)
    output.string.should == expected_output
  end

end
