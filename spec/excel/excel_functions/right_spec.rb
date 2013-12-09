require_relative '../../spec_helper.rb'

describe "ExcelFunctions: RIGHT(string,[characters])" do
  
  it "should return the right n characters from a string" do
    FunctionTest.right("ONE").should == "E"
    FunctionTest.right("ONE",1).should == "E"
    FunctionTest.right("ONE",3).should == "ONE"
  end

  it "should turn numbers into strings before processing" do
    FunctionTest.right(1.31e12,3).should == "0.0"
  end
    
  it "should turn booleans into the words TRUE and FALSE before processing" do
    FunctionTest.right(TRUE,3).should == "RUE"
    FunctionTest.right(FALSE,3).should == "LSE"
  end

  it "should return nil if given nil for either argument" do
    FunctionTest.right(nil,3).should == nil
    FunctionTest.right("ONE",nil).should == nil
  end
  
  it "should return an error if an argument is an error" do
    FunctionTest.right(:error).should == :error
    FunctionTest.right("ONE",:error).should == :error
    FunctionTest.right(:error,:error).should == :error
  end
  
  it "should be in the list of functions that can be mapped to ruby" do
    MapFormulaeToRuby::FUNCTIONS[:'RIGHT'].should == 'right'
  end
  
end
