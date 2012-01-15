require_relative '../../spec_helper.rb'
include ExcelFunctions

describe "ExcelFunctions: COSH()" do
  
  it "should calculate COSH(number)" do
    cosh(1).should == Math.cosh(1)
  end
  
end
