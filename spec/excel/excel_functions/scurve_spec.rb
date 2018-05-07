require_relative '../../spec_helper.rb'

describe "ExcelFunctions: SCURVE" do
  
  it "should return something when given appropriate arguments" do
    FunctionTest.scurve(2028,0,10,10).should == 1
  end
  
end
