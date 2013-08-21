require_relative '../spec_helper'

describe ReplaceArithmeticOnRanges do
  
  it "should replace arithmetic on ranges (e.g., 1/A1:B2) with individual calculations (i.e., [1/A1, 1/B1, 1/A2, 1/B2])" do

input = <<END
A0\t[:arithmetic, [:number, "1"], [:operator, "/"], [cell, "F198]]
A1\t[:arithmetic, [:number, "1"], [:operator, "/"], [:array, [:row, [:cell, "F197"]], [:row, [:cell, "F198"]], [:row, [:cell, "F199"]]]]
A2\t[:arithmetic, [:array, [:row, [:cell, "F197"]], [:row, [:cell, "F198"]], [:row, [:cell, "F199"]]], [:operator, "+"], [:number, "1"]]
A3\t[:arithmetic, [:array, [:row, [:cell, "F197"]], [:row, [:cell, "F198"]], [:row, [:cell, "F199"]]], [:operator, "/"],  [:array, [:row, [:cell, "F197"]], [:row, [:cell, "F198"]], [:row, [:cell, "F199"]]]]
END

expected_output = <<END
A0\t[:arithmetic, [:number, "1"], [:operator, "/"], [cell, "F198]]
A1\t[:array, [:row, [:arithmetic, [:number, "1"], [:operator, "/"], [:cell, "F197"]]], [:row, [:arithmetic, [:number, "1"], [:operator, "/"], [:cell, "F198"]]], [:row, [:arithmetic, [:number, "1"], [:operator, "/"], [:cell, "F199"]]]]
A2\t[:array, [:row, [:arithmetic, [:cell, "F197"], [:operator, "+"], [:number, "1"]]], [:row, [:arithmetic, [:cell, "F198"], [:operator, "+"], [:number, "1"]]], [:row, [:arithmetic, [:cell, "F199"], [:operator, "+"], [:number, "1"]]]]
A3\t[:array, [:row, [:arithmetic, [:cell, "F197"], [:operator, "/"], [:cell, "F197"]]], [:row, [:arithmetic, [:cell, "F198"], [:operator, "/"], [:cell, "F198"]]], [:row, [:arithmetic, [:cell, "F199"], [:operator, "/"], [:cell, "F199"]]]]
END
    
    input = StringIO.new(input)
    output = StringIO.new
    ReplaceArithmeticOnRanges.replace(input,output)
    output.string.should == expected_output
  end

end
