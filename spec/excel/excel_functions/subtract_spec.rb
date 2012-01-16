require_relative '../../spec_helper.rb'
include ExcelFunctions

describe "ExcelFunctions: subtract(number,number)" do
  
  it "should return sum of its arguments" do
    subtract(2,1).should == 1
    subtract(2.0,1.0).should == 1
  end
    
  it "should treat nil as zero" do
    subtract(1,nil).should == 1
    subtract(nil,nil).should == 0
    subtract(nil,1).should == -1
  end
  
  it "should return an error if either argument is an error" do
    subtract(:error,1).should == :error
    subtract(1,:error).should == :error
    subtract(:error1,:error2).should == :error1
  end
  
end
