require_relative '../../spec_helper.rb'

describe "ExcelFunctions: multiply(number,number)" do
  
  it "should return the multiple of its arguments" do
    FunctionTest.multiply(3,4).should == 12
    FunctionTest.multiply(0.5,1.5).should == 0.75
  end
    
  it "should treat nil as zero" do
    FunctionTest.multiply(1,nil).should == 0
    FunctionTest.multiply(nil,nil).should == 0
    FunctionTest.multiply(nil,1).should == 0
  end
  
  it "should return an error if either argument is an error" do
    FunctionTest.multiply(:error,1).should == :error
    FunctionTest.multiply(1,:error).should == :error
    FunctionTest.multiply(:error1,:error2).should == :error1
  end
  
end
