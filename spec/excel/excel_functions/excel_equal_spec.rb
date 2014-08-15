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
    FunctionTest.excel_equal?("hello","HELLO").should == true
    FunctionTest.excel_equal?("hello",1).should == false
    FunctionTest.excel_equal?(1,"HELLO").should == false
    FunctionTest.excel_equal?(nil,"HELLO").should == false
    FunctionTest.excel_equal?("hello",nil).should == false
  end

  it "should return error if either argument is an error" do
    FunctionTest.excel_equal?(:error,1).should == :error
    FunctionTest.excel_equal?(1,:error).should == :error
    FunctionTest.excel_equal?(:error,:error).should == :error
  end
  
end
