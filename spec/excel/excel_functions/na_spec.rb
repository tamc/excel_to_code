require_relative '../../spec_helper.rb'

describe "ExcelFunctions: NA" do
  
  it "should return the NA error" do
    FunctionTest.na().should == :na
  end
  
  it "should be in the list of functions that can be mapped to ruby" do
    MapFormulaeToRuby::FUNCTIONS[:'NA'].should == 'na'
  end
  
end
