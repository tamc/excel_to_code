require_relative '../../spec_helper.rb'

describe "ExcelFunctions: = not_equal?()" do
  
  it "should check the equality of numbers" do
    FunctionTest.not_equal?(1,1).should == false
    FunctionTest.not_equal?(1,0).should == true
    FunctionTest.not_equal?(1.0,1).should == false
  end
  
  it "should check the equality of booleans" do
    FunctionTest.not_equal?(true,false).should == true
    FunctionTest.not_equal?(false,false).should == false
  end

  it "should check the equality of strings, ignoring case" do
    FunctionTest.not_equal?("HELLO","world").should == true
    FunctionTest.not_equal?("HELLO","hello").should == false
  end

  # it "should be able to check arrays" do
  #   FunctionTest.not_equal?([[1,2],[3,4]],2).should == [[true,false],[true,true]]
  #   FunctionTest.not_equal?(2,[[1,2],[3,4]]).should == [[true,false],[true,true]]
  #   FunctionTest.not_equal?([[1,2],[3,4]],[[1,2],[3,4]]).should == [[false,false],[false,false]]
  # end

  it "should return error if either argument is an error" do
    FunctionTest.not_equal?(:error,1).should == :error
    FunctionTest.not_equal?(1,:error).should == :error
    FunctionTest.not_equal?(:error,:error).should == :error
  end
  
end
