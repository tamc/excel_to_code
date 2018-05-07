require_relative '../../spec_helper.rb'

describe "ExcelFunctions: FLOOR" do
  
  it "should return something when given appropriate arguments" do
    FunctionTest.floor(1990,100).should == 1900
    FunctionTest.floor(10.99,0.1).should == 10.9
  end

  it "should return an error when given inappropriate arguments" do
    FunctionTest.floor("Asdasddf", 1).should == :value
    FunctionTest.floor(1, "Aasdf").should == :value
  end
    
  it "should treat nil as zero" do
    FunctionTest.floor(nil,1).should == 0
  end
  
  it "should return an error if an argument is an error" do
    FunctionTest.floor(:error, 1).should == :error
    FunctionTest.floor(1, :error).should == :error
  end
  
  it "should be in the list of functions that can be mapped to ruby" do
    MapFormulaeToRuby::FUNCTIONS[:'FLOOR'].should == 'floor'
  end
  
end
