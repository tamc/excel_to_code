require_relative '../../spec_helper.rb'

describe "ExcelFunctions: NUMBER_OR_ZERO" do
  
  it "should return a number when given a number, an error if given an error, otherwise nothing" do
    FunctionTest.number_or_zero(0).should == 0
    FunctionTest.number_or_zero(1).should == 1
    FunctionTest.number_or_zero(:error).should == :error
    FunctionTest.number_or_zero(true).should == 0
    FunctionTest.number_or_zero(false).should == 0
    FunctionTest.number_or_zero(nil).should == 0
    FunctionTest.number_or_zero("1").should == 0
    FunctionTest.number_or_zero("Aasdfadsf").should == 0
  end
  
  it "should be in the list of functions that can be mapped to ruby" do
    MapFormulaeToRuby::FUNCTIONS[:'NUMBER_OR_ZERO'].should == 'number_or_zero'
  end
  
end
