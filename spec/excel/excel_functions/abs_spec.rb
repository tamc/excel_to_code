require_relative '../../spec_helper.rb'

describe "ExcelFunctions: abs(number,number)" do
  
  it "should return the absolute value of its argument" do
    FunctionTest.abs(1).should == 1
    FunctionTest.abs(-1).should == 1
  end
  
  it "should return a 1 if the argument is true, 0 if the argument is false" do
    FunctionTest.abs(true).should == 1
    FunctionTest.abs(false).should == 0
  end
  
  it "should return a value error if argument is a string" do
    FunctionTest.abs("Hello").should == :value
  end
    
  it "should treat nil as zero" do
    FunctionTest.abs(nil).should == 0
  end
  
  it "should return an error if the argument is an error" do
    FunctionTest.abs(:error).should == :error
  end
  
end
