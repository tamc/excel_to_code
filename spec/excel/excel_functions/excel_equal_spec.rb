require_relative '../../spec_helper.rb'

describe "ExcelFunctions: = excel_equal?()" do
  
  it "should check the equality of numbers" do
    FunctionTest.excel_equal?(1,1).should == true
    FunctionTest.excel_equal?(1,0).should == false
    FunctionTest.excel_equal?(1.0,1).should == true
  end
  
  it "should check the equality of booleans" do
    FunctionTest.excel_equal?(true,false).should == false
    FunctionTest.excel_equal?(false,false).should == true
  end

  it "should check the equality of strings, ignoring case" do
    FunctionTest.excel_equal?("HELLO","world").should == false
    FunctionTest.excel_equal?("HELLO","hello").should == true
  end
  
  it "should be able to check arrays" do
    FunctionTest.excel_equal?([[1,2],[3,4]],2).should == [[false,true],[false,false]]
    FunctionTest.excel_equal?(2,[[1,2],[3,4]]).should == [[false,true],[false,false]]
    FunctionTest.excel_equal?([[1,2],[3,4]],[[1,2],[3,4]]).should == [[true,true],[true,true]]
  end

  it "should return error if either argument is an error" do
    FunctionTest.excel_equal?(:error,1).should == :error
    FunctionTest.excel_equal?(1,:error).should == :error
    FunctionTest.excel_equal?(:error,:error).should == :error
  end
  
end
