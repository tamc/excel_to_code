require_relative '../spec_helper'

describe ReplaceArraysWithSingleCells do
  
  it "should replace array literals (e.g., {A1,B1;A2,B2}) with the first cell (e.g., A1) where it is the only thing in the formula" do

input = <<END
A1\t[:array, [:row, [:sheet_reference, "sheet1", [:cell, "A1"]], [:sheet_reference, "sheet1", [:cell, "B1"]]], [:row, [:sheet_reference, "sheet1", [:cell, "A2"]], [:sheet_reference, "sheet1", [:cell, "B2"]]]]
A2\t[:funtion, "SUM", [:array, [:row, [:sheet_reference, "sheet1", [:cell, "A1"]], [:sheet_reference, "sheet1", [:cell, "B1"]]], [:row, [:sheet_reference, "sheet1", [:cell, "A2"]], [:sheet_reference, "sheet1", [:cell, "B2"]]]]]
END

expected_output = <<END
A1\t[:sheet_reference, "sheet1", [:cell, "A1"]]
A2\t[:funtion, "SUM", [:array, [:row, [:sheet_reference, "sheet1", [:cell, "A1"]], [:sheet_reference, "sheet1", [:cell, "B1"]]], [:row, [:sheet_reference, "sheet1", [:cell, "A2"]], [:sheet_reference, "sheet1", [:cell, "B2"]]]]]
END
    
    input = StringIO.new(input)
    output = StringIO.new
    ReplaceArraysWithSingleCells.replace(input,output)
    output.string.should == expected_output
  end

end
