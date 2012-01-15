require_relative '../../spec_helper.rb'
include ExcelFunctions

describe "ExcelFunctions: SUM()" do
  
  it "should return the sum of its arguments" do
    sum(1,2,3).should == 6
  end

  it "should return the sum of its arguments, flattening arrays" do
    sum([[1],[2],[3]]).should == 6
  end
  
  it "should ignore any arguments that are not numbers" do
    sum(1,true,2,"Hello",3).should == 6
  end
  
end
