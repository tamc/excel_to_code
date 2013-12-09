require_relative '../../spec_helper.rb'

describe "ExcelFunctions: PI()" do
  
  it "should return the value of PI" do
    FunctionTest.pi().should == Math::PI
  end
  
end
