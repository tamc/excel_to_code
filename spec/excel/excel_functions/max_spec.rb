require_relative '../../spec_helper.rb'

describe "ExcelFunctions: MAX" do
  
  it "should return the argument with the greatest value, flattening arrays" do
    FunctionTest.max(1,[10,100]).should == 100    
  end
  
  it "should ignore non numeric values" do
    FunctionTest.max("Asdasddf","aasdfa",1).should == 1
  end
    
  it "should ignore nil values" do
    FunctionTest.max(-10,nil).should == -10
  end
  
  it "should return zero if no arguments" do
    FunctionTest.max(nil).should == 0
  end
  
  it "should return an error if an argument is an error" do
    FunctionTest.max(:error).should == :error
    FunctionTest.max([[:na,:na]]).should == :na
  end
  
  it "should be in the list of functions that can be mapped to ruby" do
    MapFormulaeToRuby::FUNCTIONS[:'MAX'].should == 'max'
  end
  
end
