require_relative '../../spec_helper.rb'

describe "ExcelFunctions: iserrorOR" do
  
  it "should return true if passed any other sort of error" do
    FunctionTest.iserror(:na).should == true
    FunctionTest.iserror(:div0).should == true
    FunctionTest.iserror(:ref).should == true
    FunctionTest.iserror(:value).should == true
    FunctionTest.iserror(:name).should == true
  end
    
  it "should return false in all other cases" do
    FunctionTest.iserror(nil).should == false
    FunctionTest.iserror(true).should == false
    FunctionTest.iserror(false).should == false
    FunctionTest.iserror(10.0).should == false
    FunctionTest.iserror("Hello").should == false
  end
  
  it "should be in the list of functions that can be mapped to ruby" do
    MapFormulaeToRuby::FUNCTIONS[:'ISERROR'].should == 'iserror'
  end
  
end
