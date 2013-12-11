require_relative '../spec_helper'

describe ReplaceRangesWithArrayLiterals do
  
  it "should replace ranges (e.g., A1:B2) with array literals (e.g., {A1,B1;A2,B2}) except if they only refer to a single cell, in which case replace with a single cell" do

input = <<END
A1\t[:sheet_reference, :sheet1, [:area, :A1, :B2]]
A2\t[:area, :"F$197", :"F$199"]
A3\t[:area, :"$F197", :"F$197"]
END

expected_output = <<END
A1\t[:array, [:row, [:sheet_reference, :sheet1, [:cell, :A1]], [:sheet_reference, :sheet1, [:cell, :B1]]], [:row, [:sheet_reference, :sheet1, [:cell, :A2]], [:sheet_reference, :sheet1, [:cell, :B2]]]]
A2\t[:array, [:row, [:cell, :F197]], [:row, [:cell, :F198]], [:row, [:cell, :F199]]]
A3\t[:cell, :F197]
END
    
    input = StringIO.new(input)
    output = StringIO.new
    ReplaceRangesWithArrayLiterals.replace(input,output)
    output.string.should == expected_output
  end

  it "should return exactly the same array literal for every sheet reference expansion" do
    testA = [:sheet_reference, :sheet1, [:area, :A1, :B2]]
    testB = [:sheet_reference, :sheet1, [:area, :A1, :B2]]
    r = ReplaceRangesWithArrayLiteralsAst.new
    first = r.map(testA)
    second = r.map(testB)
    first.object_id.should == second.object_id
  end

end
