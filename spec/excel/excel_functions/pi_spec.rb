require_relative '../../spec_helper.rb'
include ExcelFunctions

describe "ExcelFunctions: PI()" do
  
  it "should return the value of PI" do
    pi().should == Math::PI
  end
  
end