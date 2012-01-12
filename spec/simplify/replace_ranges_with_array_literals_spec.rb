require_relative '../spec_helper'

describe ReplaceRangesWithArrayLiterals do
  
  it "should replace ranges (e.g., A1:B2) with array literals (e.g., {A1,B1;A2,B2}) " do

input = <<END
A1\t[:sheet_reference, "sheet1", [:area, "A1", "B2"]]
END

expected_output = <<END
A1\t[:array, [:row, [:sheet_reference, "sheet1", [:cell, "A1"]], [:sheet_reference, "sheet1", [:cell, "B1"]]], [:row, [:sheet_reference, "sheet1", [:cell, "A2"]], [:sheet_reference, "sheet1", [:cell, "B2"]]]]
END
    
    input = StringIO.new(input)
    output = StringIO.new
    ReplaceRangesWithArrayLiterals.replace(input,output)
    output.string.should == expected_output
  end

end
