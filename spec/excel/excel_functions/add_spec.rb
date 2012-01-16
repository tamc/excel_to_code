require_relative '../../spec_helper.rb'
include ExcelFunctions

describe "ExcelFunctions: add(number,number)" do
  
  it "should return sum of its arguments" do
    add(1,1).should == 2
    add(1.0,1.0).should == 2.0
  end
    
  it "should treat nil as zero" do
    add(1,nil).should == 1
    add(nil,nil).should == 0
    add(nil,1).should == 1
  end
  
  it "should return an error if either argument is an error" do
    add(:error,1).should == :error
    add(1,:error).should == :error
    add(:error1,:error2).should == :error1
  end
  
end
