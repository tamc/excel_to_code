require_relative '../../spec_helper.rb'
include ExcelFunctions

describe "ExcelFunctions: = more_than_or_equal?()" do
  
  it "should check the equality of numbers" do
    more_than_or_equal?(1,2).should == false
    more_than_or_equal?(1,1).should == true
    more_than_or_equal?(1,0).should == true
    more_than_or_equal?(1.0,1).should == true
  end
  
  it "should check the equality of booleans" do
    more_than_or_equal?(true,true).should == true
    more_than_or_equal?(true,false).should == true
    more_than_or_equal?(false,true).should == false
    more_than_or_equal?(false,false).should == true
  end

  it "should check the equality of strings, ignoring case" do
    more_than_or_equal?("HELLO","world").should == false
    more_than_or_equal?("HELLO","hello").should == true
  end

  it "should return error if either argument is an error" do
    more_than_or_equal?(:error,1).should == :error
    more_than_or_equal?(1,:error).should == :error
    more_than_or_equal?(:error,:error).should == :error
  end
  
end
