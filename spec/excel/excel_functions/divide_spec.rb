require_relative '../../spec_helper.rb'
include ExcelFunctions

describe "ExcelFunctions: divide(number,number)" do
  
  it "should return sum of its arguments" do
    divide(4,2).should == 2
    divide(1.0,0.5).should == 2.0
  end
  
  it "should return a :div0! if a divide by zero occurs" do
    divide(1,0).should == :div0
  end
    
  it "should treat nil as zero" do
    divide(1,nil).should == :div0
    divide(nil,nil).should == :div0
    divide(nil,1).should == 0
  end
  
  it "should return an error if either argument is an error" do
    divide(:error,1).should == :error
    divide(1,:error).should == :error
    divide(:error1,:error2).should == :error1
  end
  
end
