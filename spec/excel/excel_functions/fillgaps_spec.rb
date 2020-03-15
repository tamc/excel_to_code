require_relative '../../spec_helper.rb'

describe "ExcelFunctions: FILLGAPS" do
  
  it "should return something when given appropriate arguments" do
    FunctionTest.fillgaps([[1]], [[2001]], 2001).should ==[[1]] 
  end
  
  it "should be in the list of functions that can be mapped to ruby" do
    MapFormulaeToRuby::FUNCTIONS[:'FILLGAPS'].should == 'fillgaps'
  end
  
end
