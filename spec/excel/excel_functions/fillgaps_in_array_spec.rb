require_relative '../../spec_helper.rb'

describe "ExcelFunctions: FILLGAPS_IN_ARRAY" do
  
  it "Should fill in blank values in passed array" do
    FunctionTest.fillgaps_in_array(4,1,[[1,nil,nil,4]], [[2001, 2002, 2003, 2004]], 2004).should ==[[1,2,3,4]] 
  end

  it "Should extrapolate blank values at the end of the array" do
    FunctionTest.fillgaps_in_array(4,1,[[2,3,4,nil]], [[2001, 2002, 2003, 2004]], 2004).should ==[[2,3,4,5]] 

  end

  
  it "should be in the list of functions that can be mapped to ruby" do
    MapFormulaeToRuby::FUNCTIONS[:'FILLGAPS_IN_ARRAY'].should == 'fillgaps_in_array'
  end
  
end
