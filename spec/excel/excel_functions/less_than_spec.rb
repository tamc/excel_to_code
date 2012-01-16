require_relative '../../spec_helper.rb'
include ExcelFunctions

describe "ExcelFunctions: = less_than?()" do
  
  it "should check the equality of numbers" do
    less_than?(1,2).should == true
    less_than?(1,1).should == false
    less_than?(1,0).should == false
    less_than?(1.0,1).should == false
  end
  
  it "should check the equality of booleans" do
    less_than?(true,true).should == false
    less_than?(true,false).should == false
    less_than?(false,true).should == true
    less_than?(false,false).should == false
  end

  it "should check the equality of strings, ignoring case" do
    less_than?("HELLO","world").should == true
    less_than?("HELLO","hello").should == false
  end

  it "should return error if either argument is an error" do
    less_than?(:error,1).should == :error
    less_than?(1,:error).should == :error
    less_than?(:error,:error).should == :error
  end
  
end
