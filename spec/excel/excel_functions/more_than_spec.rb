require_relative '../../spec_helper.rb'

describe "ExcelFunctions: = more_than?()" do
  
  it "should check the equality of numbers" do
    FunctionTest.more_than?(1,2).should == false
    FunctionTest.more_than?(1,1).should == false
    FunctionTest.more_than?(1,0).should == true
    FunctionTest.more_than?(1.0,1).should == false
  end

  it "should check the equality of inputs of different types" do
    FunctionTest.more_than?(true,"hello").should == true
    FunctionTest.more_than?("hello", 1).should == true
    FunctionTest.more_than?(1,"hello").should == false
    FunctionTest.more_than?("hello", true).should == false
  end
  
  it "should check the equality of booleans" do
    FunctionTest.more_than?(false,false).should == false
    FunctionTest.more_than?(false,true).should == false
    FunctionTest.more_than?(true,false).should == true
    FunctionTest.more_than?(true,true).should == false
  end

  it "should check the equality of strings, ignoring case" do
    FunctionTest.more_than?("HELLO","world").should == false
    FunctionTest.more_than?("HELLO","hello").should == false
  end
  
  it "should treat nil values as zero" do
    FunctionTest.more_than?(nil,1).should == false
    FunctionTest.more_than?(nil,-1).should == true
    FunctionTest.more_than?(1,nil).should == true
    FunctionTest.more_than?(-1,nil).should == false
  end

  # it "should be able to check arrays" do
  #   FunctionTest.more_than?([[1,2],[3,4]],2).should == [[false,false],[true,true]]
  #   FunctionTest.more_than?(2,[[1,2],[3,4]]).should == [[true,false],[false,false]]
  #   FunctionTest.more_than?([[1,2],[3,4]],[[1,2],[3,4]]).should == [[false,false],[false,false]]
  # end

  it "should return error if either argument is an error" do
    FunctionTest.more_than?(:error,1).should == :error
    FunctionTest.more_than?(1,:error).should == :error
    FunctionTest.more_than?(:error,:error).should == :error
  end
  
end
