require_relative '../../spec_helper.rb'

describe "ExcelFunctions: LEN" do
  
  it "should return the length of a passed string, if not a string, convert to a string first" do
    FunctionTest.len(nil).should == 0
    FunctionTest.len("Hello").should == 5
    FunctionTest.len(123).should == 3
    FunctionTest.len(true).should == 4
    FunctionTest.len(false).should == 5
  end

  it "should return an error if an argument is an error" do
    FunctionTest.len(:error).should == :error
  end
  
  it "should be in the list of functions that can be mapped to ruby" do
    MapFormulaeToRuby::FUNCTIONS[:'LEN'].should == 'len'
  end
  
end
