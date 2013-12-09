require_relative '../../spec_helper.rb'

describe "ExcelFunctions: LEFT(string,[characters])" do
  
  it "should return the left n characters from a string" do
    FunctionTest.left("ONE").should == "O"
    FunctionTest.left("ONE",1).should == "O"
    FunctionTest.left("ONE",3).should == "ONE"
  end

  it "should turn numbers into strings before processing" do
    FunctionTest.left(1.31e12,3).should == "131"
  end
    
  it "should turn booleans into the words TRUE and FALSE before processing" do
    FunctionTest.left(TRUE,3).should == "TRU"
    FunctionTest.left(FALSE,3).should == "FAL"
  end

  it "should return nil if given nil for either argument" do
    FunctionTest.left(nil,3).should == nil
    FunctionTest.left("ONE",nil).should == nil
  end
  
  it "should return an error if an argument is an error" do
    FunctionTest.left(:error).should == :error
    FunctionTest.left("ONE",:error).should == :error
    FunctionTest.left(:error,:error).should == :error
  end
  
  it "should be in the list of functions that can be mapped to ruby" do
    MapFormulaeToRuby::FUNCTIONS[:'LEFT'].should == 'left'
  end
  
end
