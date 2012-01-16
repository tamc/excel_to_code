require_relative '../../spec_helper.rb'
include ExcelFunctions

describe "ExcelFunctions: = less_than_or_equal?()" do
  
  it "should check the equality of numbers" do
    less_than_or_equal?(1,2).should == true
    less_than_or_equal?(1,1).should == true
    less_than_or_equal?(1,0).should == false
    less_than_or_equal?(1.0,1).should == true
  end
  
  it "should check the equality of booleans" do
    less_than_or_equal?(true,true).should == true
    less_than_or_equal?(true,false).should == false
    less_than_or_equal?(false,true).should == true
    less_than_or_equal?(false,false).should == true
  end

  it "should check the equality of strings, ignoring case" do
    less_than_or_equal?("HELLO","world").should == true
    less_than_or_equal?("HELLO","hello").should == true
  end

  it "should return error if either argument is an error" do
    less_than_or_equal?(:error,1).should == :error
    less_than_or_equal?(1,:error).should == :error
    less_than_or_equal?(:error,:error).should == :error
  end
  
end
