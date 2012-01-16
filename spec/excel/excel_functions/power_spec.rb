require_relative '../../spec_helper.rb'
include ExcelFunctions

describe "ExcelFunctions: power(number,number)" do
  
  it "should return sum of its arguments" do
    power(2,3).should == 8
    power(4.0,0.5).should == 2.0
  end
      
  it "should treat nil as zero" do
    power(1,nil).should == 1
    power(nil,nil).should == 1
    power(nil,1).should == 0
  end
  
  it "should return an error if either argument is an error" do
    power(:error,1).should == :error
    power(1,:error).should == :error
    power(:error1,:error2).should == :error1
  end
  
end
