require_relative '../spec_helper'

describe ReplaceStringJoinOnRanges do
  
  it "should replace string joins on ranges (e.g., A1:A2&B1:B2 = [A1&B1, A2&B2]). Note Excel only seems to do this if formulae are entered in array form?" do

input = <<END
A0\t[:string_join, [:string, "A"], [:number, 1]
A1\t[:string_join, [:string, "A"], [:array, [:row, [:number, 1]], [:row, [:number, 2]]]]
A2\t[:string_join, [:array, [:row, [:string, "A"]], [:row, [:string, "B"]]], [:array, [:row, [:number, 1]], [:row, [:number, 2]]]]
A3\t[:string_join, [:array, [:row, [:string, "A"]], [:row, [:string, "B"]]], [:number, 1]]
END

expected_output = <<END
A0\t[:string_join, [:string, "A"], [:number, 1]
A1\t[:array, [:row, [:string_join, [:string, "A"], [:number, 1]]], [:row, [:string_join, [:string, "A"], [:number, 2]]]]
A2\t[:array, [:row, [:string_join, [:string, "A"], [:number, 1]]], [:row, [:string_join, [:string, "B"], [:number, 2]]]]
A3\t[:array, [:row, [:string_join, [:string, "A"], [:number, 1]]], [:row, [:string_join, [:string, "B"], [:number, 1]]]]
END
    
    input = StringIO.new(input)
    output = StringIO.new
    ReplaceStringJoinOnRanges.replace(input,output)
    output.string.should == expected_output
  end

end
