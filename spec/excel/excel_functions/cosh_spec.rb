require_relative '../../spec_helper.rb'

describe "ExcelFunctions: COSH()" do
  
  it "should calculate COSH(number)" do
    FunctionTest.cosh(1).should == Math.cosh(1)
  end
    
  it "should treat nil as zero" do
    FunctionTest.cosh(nil).should == Math.cosh(0)
  end
  
  it "should return error if argument is an error" do
    FunctionTest.cosh(:error).should == :error
  end
  
end
