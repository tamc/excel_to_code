require_relative '../../spec_helper.rb'

# ADDRESS(row_num,column_num,abs_num,a1,sheet_text)
# row_num
# The row number to use in the cell reference.
# column_num
# The column number to use in the cell reference.
# abs_num
# Specifies the type of reference to return.
# 1 or omitted returns the following type of reference: absolute.
# 2 returns the following type of reference: absolute row; relative column.
# 3 returns the following type of reference: relative row; absolute column.
# 4 returns the following type of reference: relative.
# a1
# A logical value that specifies the A1 or R1C1 reference style.
# If a1 is TRUE or omitted, ADDRESS returns an A1-style reference.
# If a1 is FALSE, ADDRESS returns an R1C1-style reference.
# sheet_text
# Text specifying the name of the sheet to be used as the external reference.
# If sheet_text is omitted, no sheet name is used.

describe "ExcelFunctions: ADDRESS" do
  
  it "should return a reference as a string given a row and a column number" do
    FunctionTest.address(1,1).should == "$A$1"
    FunctionTest.address(2,2).should == "$B$2"
    FunctionTest.address(1,1,1).should == "$A$1"
    FunctionTest.address(1,1,2).should == "A$1"
    FunctionTest.address(1,1,3).should == "$A1"
    FunctionTest.address(1,1,4).should == "A1"
    FunctionTest.address(1,1,4,true).should == "A1"
    FunctionTest.address(1,1,4,true,"sheet1").should == "'sheet1'!A1"
    FunctionTest.address(1.7,3.9).should == "$C$1"
  end

  it "should return an error when given inappropriate arguments" do
    FunctionTest.address(0,1).should == :value
    FunctionTest.address(1,0).should == :value
  end
   
  it "should return an error if an argument is an error" do
    FunctionTest.address(1, :error).should == :error
  end
  
  it "should be in the list of functions that can be mapped to ruby" do
    MapFormulaeToRuby::FUNCTIONS[:'ADDRESS'].should == 'address'
  end
  
end
