require_relative '../../spec_helper.rb'

describe "ExcelFunctions: ISERR" do
  
  it "should return false if the passed error is #N/A (wtf?)" do
    FunctionTest.iserr(:na).should == false
  end

  it "should return true if passed any other sort of error" do
    FunctionTest.iserr(:div0).should == true
    FunctionTest.iserr(:ref).should == true
    FunctionTest.iserr(:value).should == true
    FunctionTest.iserr(:name).should == true
  end
    
  it "should return false in all other cases" do
    FunctionTest.iserr(nil).should == false
    FunctionTest.iserr(true).should == false
    FunctionTest.iserr(false).should == false
    FunctionTest.iserr(10.0).should == false
    FunctionTest.iserr("Hello").should == false
  end
  
  it "should be in the list of functions that can be mapped to ruby" do
    MapFormulaeToRuby::FUNCTIONS[:'ISERR'].should == 'iserr'
  end
  
end
