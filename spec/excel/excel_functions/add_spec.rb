require_relative '../../spec_helper.rb'
include ExcelFunctions

describe "ExcelFunctions: add(number,number)" do
  
  it "should return sum of its arguments" do
    add(1,1).should == 2
    add(1.0,1.0).should == 2.0
  end
  
end
