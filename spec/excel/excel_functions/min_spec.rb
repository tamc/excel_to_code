require_relative '../../spec_helper.rb'

describe "ExcelFunctions: MIN" do
  
  it "should return the argument with the smallest value, flattening arrays" do
    FunctionTest.min(1000,[10,100]).should == 10
  end
  
  it "should ignore non numeric values" do
    FunctionTest.min("Asdasddf","aasdfa",1).should == 1
  end
    
  it "should ignore nil values" do
    FunctionTest.min(10,nil).should == 10
  end
  
  it "should return zero if no arguments" do
    FunctionTest.min(nil).should == 0
  end
  
  it "should return an error if an argument is an error" do
    FunctionTest.min(:error).should == :error
  end

  
  it "should be in the list of functions that can be mapped to ruby" do
    MapFormulaeToRuby::FUNCTIONS[:'MIN'].should == 'min'
  end
  
end
