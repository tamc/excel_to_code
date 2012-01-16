require_relative '../../spec_helper.rb'
include ExcelFunctions

describe "ExcelFunctions: = not_equal?()" do
  
  it "should check the equality of numbers" do
    not_equal?(1,1).should == false
    not_equal?(1,0).should == true
    not_equal?(1.0,1).should == false
  end
  
  it "should check the equality of booleans" do
    not_equal?(true,false).should == true
    not_equal?(false,false).should == false
  end

  it "should check the equality of strings, ignoring case" do
    not_equal?("HELLO","world").should == true
    not_equal?("HELLO","hello").should == false
  end

  it "should return error if either argument is an error" do
    not_equal?(:error,1).should == :error
    not_equal?(1,:error).should == :error
    not_equal?(:error,:error).should == :error
  end
  
end
