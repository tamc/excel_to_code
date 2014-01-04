require_relative '../../spec_helper.rb'

describe "ExcelFunctions: LOWER" do
  
  it "should return a lower case string of whatever is passed to it" do
    FunctionTest.lower("HELLO").should == "hello"
    FunctionTest.lower(1).should == "1"
    FunctionTest.lower(true).should == "true"
    FunctionTest.lower(false).should == "false"
  end

  it "should treat nil as an empty string" do
    FunctionTest.lower(nil).should == "" 
  end
  
  it "should return an error if an argument is an error" do
    FunctionTest.lower(:error).should == :error
  end
  
  it "should be in the list of functions that can be mapped to ruby" do
    MapFormulaeToRuby::FUNCTIONS['LOWER'].should == 'lower'
  end
  
end
