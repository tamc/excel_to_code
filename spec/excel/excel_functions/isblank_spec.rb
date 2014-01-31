require_relative '../../spec_helper.rb'

describe "ExcelFunctions: ISBLANK" do
  
  it "should return true if passed an empty value and false in all other situations" do
    FunctionTest.isblank(nil).should == true
    FunctionTest.isblank(1).should == false
    FunctionTest.isblank(:error).should == false
    FunctionTest.isblank(true).should == false
    FunctionTest.isblank("").should == false
  end

  
  it "should be in the list of functions that can be mapped to ruby" do
    MapFormulaeToRuby::FUNCTIONS[:'ISBLANK'].should == 'isblank'
  end
  
end
