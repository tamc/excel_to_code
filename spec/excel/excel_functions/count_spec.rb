require_relative '../../spec_helper.rb'

describe "ExcelFunctions: COUNT(arg1,[arg2],..)" do
  
  it "should count the number of numeric values in an area" do
    FunctionTest.count(1,"two",[[10],[100],[:error],[nil]]).should == 3    
  end

  it "should not count strings that contain numbers" do
    FunctionTest.count(1,"2",[[10],[100],[nil]]).should == 3    
  end
    
  it "should not count nil" do
    FunctionTest.count(nil).should == 0
  end
  
  it "should ignore errors" do
    FunctionTest.count(10,:error).should == 1
    FunctionTest.count(:error).should == 0
  end
  
  it "should be in the list of functions that can be mapped to ruby" do
    MapFormulaeToRuby::FUNCTIONS[:'COUNT'].should == 'count'
  end
  
end
