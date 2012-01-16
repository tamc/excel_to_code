require_relative '../../spec_helper.rb'
include ExcelFunctions

describe "ExcelFunctions: COSH()" do
  
  it "should calculate COSH(number)" do
    cosh(1).should == Math.cosh(1)
  end
    
  it "should treat nil as zero" do
    cosh(nil).should == Math.cosh(0)
  end
  
  it "should return error if argument is an error" do
    cosh(:error).should == :error
  end
  
end
