require_relative '../../spec_helper.rb'

describe "ExcelFunctions: VALUE" do
  
  it "should return a number when passed a number" do
    FunctionTest.value(1).should == 1
  end

  it "should return a number when passed a string containing a number" do
    FunctionTest.value("1").should == 1
    FunctionTest.value("1.01").should == 1.01
    FunctionTest.value("1.01e3").should == 1010
  end

  it "should return a value error when given something that isn't a string" do
    FunctionTest.value("Asdasddf").should == :value
  end
    
  it "should treat nil as zero" do
    FunctionTest.value(nil).should == 0
  end
  
  it "should return an error if an argument is an error" do
    FunctionTest.value(:error).should == :error
  end
  
  it "should be in the list of functions that can be mapped to ruby" do
    MapFormulaeToRuby::FUNCTIONS[:'VALUE'].should == 'value'
  end
  
end
