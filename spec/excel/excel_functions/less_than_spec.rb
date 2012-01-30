require_relative '../../spec_helper.rb'

describe "ExcelFunctions: = less_than?()" do
  
  it "should check the equality of numbers" do
    FunctionTest.less_than?(1,2).should == true
    FunctionTest.less_than?(1,1).should == false
    FunctionTest.less_than?(1,0).should == false
    FunctionTest.less_than?(1.0,1).should == false
  end
  
  it "should check the equality of booleans" do
    FunctionTest.less_than?(true,true).should == false
    FunctionTest.less_than?(true,false).should == false
    FunctionTest.less_than?(false,true).should == true
    FunctionTest.less_than?(false,false).should == false
  end

  it "should check the equality of strings, ignoring case" do
    FunctionTest.less_than?("HELLO","world").should == true
    FunctionTest.less_than?("HELLO","hello").should == false
  end

  it "should treat nil values as zero" do
    FunctionTest.less_than?(nil,1).should == true
    FunctionTest.less_than?(nil,-1).should == false
    FunctionTest.less_than?(1,nil).should == false
    FunctionTest.less_than?(-1,nil).should == true
  end

  it "should return error if either argument is an error" do
    FunctionTest.less_than?(:error,1).should == :error
    FunctionTest.less_than?(1,:error).should == :error
    FunctionTest.less_than?(:error,:error).should == :error
  end
  
end
