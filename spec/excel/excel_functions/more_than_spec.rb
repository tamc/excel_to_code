require_relative '../../spec_helper.rb'
include ExcelFunctions

describe "ExcelFunctions: = more_than?()" do
  
  it "should check the equality of numbers" do
    more_than?(1,2).should == false
    more_than?(1,1).should == false
    more_than?(1,0).should == true
    more_than?(1.0,1).should == false
  end
  
  it "should check the equality of booleans" do
    more_than?(false,false).should == false
    more_than?(false,true).should == false
    more_than?(true,false).should == true
    more_than?(true,true).should == false
  end

  it "should check the equality of strings, ignoring case" do
    more_than?("HELLO","world").should == false
    more_than?("HELLO","hello").should == false
  end

  it "should return error if either argument is an error" do
    more_than?(:error,1).should == :error
    more_than?(1,:error).should == :error
    more_than?(:error,:error).should == :error
  end
  
end
