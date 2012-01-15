require_relative '../../spec_helper.rb'
include ExcelFunctions

describe "ExcelFunctions: multiply(number,number)" do
  
  it "should return the multiple of its arguments" do
    multiply(3,4).should == 12
    multiply(0.5,1.5).should == 0.75
  end
  
end
