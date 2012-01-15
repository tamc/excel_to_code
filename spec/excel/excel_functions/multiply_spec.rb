require_relative '../../spec_helper.rb'
include ExcelFunctions

describe "ExcelFunctions: multiply(number,number)" do
  
  it "should return the multiple of its arguments" do
    multiply(3,4).should == 12
    multiply(0.5,1.5).should == 0.75
  end
    
  it "should treat nil as zero" do
    multiply(1,nil).should == 0
    multiply(nil,nil).should == 0
    multiply(nil,1).should == 0
  end
  
  
end
