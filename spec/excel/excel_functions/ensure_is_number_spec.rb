require_relative '../../spec_helper.rb'

describe "ExcelFunctions: ENSURE_IS_NUMBER" do
  
  it "should ensure the passed argument becomes a number where possible" do
    FunctionTest.ensure_is_number(1).should == 1
    FunctionTest.ensure_is_number(nil).should == 0
    FunctionTest.ensure_is_number(true).should == 1
    FunctionTest.ensure_is_number(false).should == 0
    FunctionTest.ensure_is_number("1.0").should == 1.0
  end

  it "should return an error when given inappropriate arguments" do
    FunctionTest.ensure_is_number("Asdasddf").should == :value
  end
    
  it "should return an error if an argument is an error" do
    FunctionTest.ensure_is_number(:error).should == :error
  end
  
  it "should be in the list of functions that can be mapped to ruby" do
    MapFormulaeToRuby::FUNCTIONS[:'ENSURE_IS_NUMBER'].should == 'ensure_is_number'
  end
  
end
