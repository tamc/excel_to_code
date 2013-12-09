require_relative '../../spec_helper.rb'

describe "ExcelFunctions: ISNUMBER" do
  
  it "should return true if a number, false otherwise" do
    FunctionTest.isnumber(1).should == true
    FunctionTest.isnumber(true).should == false
    FunctionTest.isnumber("Hello").should == false
    FunctionTest.isnumber(nil).should == false
    FunctionTest.isnumber(:ref).should == false
    FunctionTest.isnumber([1,1]).should == false
  end

  it "should be in the list of functions that can be mapped to ruby" do
    MapFormulaeToRuby::FUNCTIONS[:'ISNUMBER'].should == 'isnumber'
  end
  
end
