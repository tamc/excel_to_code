require_relative '../../spec_helper.rb'
include ExcelFunctions

describe "ExcelFunctions: = excel_equal?()" do
  
  it "should check the equality of numbers" do
    excel_equal?(1,1).should == true
    excel_equal?(1,0).should == false
    excel_equal?(1.0,1).should == true
  end
  
  it "should check the equality of booleans" do
    excel_equal?(true,false).should == false
    excel_equal?(false,false).should == true
  end

  it "should check the equality of strings, ignoring case" do
    excel_equal?("HELLO","world").should == false
    excel_equal?("HELLO","hello").should == true
  end

  it "should return error if either argument is an error" do
    excel_equal?(:error,1).should == :error
    excel_equal?(1,:error).should == :error
    excel_equal?(:error,:error).should == :error
  end
  
end
