require_relative '../../spec_helper.rb'

describe "ExcelFunctions: RATE" do
  
  it "should return the cagr when second argument is zero" do
    (FunctionTest.rate(12,0,-69999,64786) * 1000.0).round.should == -6
  end

  
  it "should be in the list of functions that can be mapped to ruby" do
    MapFormulaeToRuby::FUNCTIONS[:'RATE'].should == 'rate'
  end
  
end
