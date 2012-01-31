require_relative '../../spec_helper.rb'

describe "ExcelFunctions: = less_than_or_equal?()" do
  
  it "should check the equality of numbers" do
    FunctionTest.less_than_or_equal?(1,2).should == true
    FunctionTest.less_than_or_equal?(1,1).should == true
    FunctionTest.less_than_or_equal?(1,0).should == false
    FunctionTest.less_than_or_equal?(1.0,1).should == true
  end
  
  it "should check the equality of booleans" do
    FunctionTest.less_than_or_equal?(true,true).should == true
    FunctionTest.less_than_or_equal?(true,false).should == false
    FunctionTest.less_than_or_equal?(false,true).should == true
    FunctionTest.less_than_or_equal?(false,false).should == true
  end

  it "should check the equality of strings, ignoring case" do
    FunctionTest.less_than_or_equal?("HELLO","world").should == true
    FunctionTest.less_than_or_equal?("HELLO","hello").should == true
  end
  
  it "should treat nil values as zero" do
    FunctionTest.less_than_or_equal?(nil,0).should == true
    FunctionTest.less_than_or_equal?(nil,1).should == true
    FunctionTest.less_than_or_equal?(nil,-1).should == false
    FunctionTest.less_than_or_equal?(1,nil).should == false
    FunctionTest.less_than_or_equal?(-1,nil).should == true
  end

  it "should be able to check arrays" do
    FunctionTest.less_than_or_equal?([[1,2],[3,4]],2).should == [[true,true],[false,false]]
    FunctionTest.less_than_or_equal?(2,[[1,2],[3,4]]).should == [[false,true],[true,true]]
    FunctionTest.less_than_or_equal?([[1,2],[3,4]],[[1,2],[3,4]]).should == [[true,true],[true,true]]
  end
  
  it "should return error if either argument is an error" do
    FunctionTest.less_than_or_equal?(:error,1).should == :error
    FunctionTest.less_than_or_equal?(1,:error).should == :error
    FunctionTest.less_than_or_equal?(:error,:error).should == :error
  end
  
end
