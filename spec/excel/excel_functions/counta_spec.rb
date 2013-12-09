require_relative '../../spec_helper.rb'

describe "ExcelFunctions: COUNTA(arg1,[arg2],..)" do
  
  it "should count the number of non blank arguments" do
    FunctionTest.counta(1,"two",[[10],[100],[:error],[nil]]).should == 5
  end
    
  it "should not count nil" do
    FunctionTest.counta(nil).should == 0
  end
  
  it "should count errors" do
    FunctionTest.counta(:error).should == 1
  end
  
  it "should be in the list of functions that can be mapped to ruby" do
    MapFormulaeToRuby::FUNCTIONS[:'COUNTA'].should == 'counta'
  end
  
end
