require_relative '../../spec_helper.rb'

describe "ExcelFunctions: CHAR" do
  
  it "should return the character for a given windows character number (annoyingly, probably different on a mac)" do
    FunctionTest.char(0).should == :value
    FunctionTest.char(256).should == :value
    FunctionTest.char(1).should == "\x01"
    FunctionTest.char(97).should == "a"
    FunctionTest.char(97.123).should == "a"
  end

  it "should return an error when given inappropriate arguments" do
    FunctionTest.char("Asdasddf").should == :value
  end
    
  it "should treat nil as zero" do
    FunctionTest.char(nil).should == :value
  end
  
  it "should return an error if an argument is an error" do
    FunctionTest.char(:error).should == :error
  end
  
  it "should be in the list of functions that can be mapped to ruby" do
    MapFormulaeToRuby::FUNCTIONS[:'CHAR'].should == 'char'
  end
  
end
