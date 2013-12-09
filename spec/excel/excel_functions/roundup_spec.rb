require_relative '../../spec_helper.rb'

describe "ExcelFunctions: ROUNDUP(number,decimal places)" do
  
  it "should round numbers up correctly" do
    FunctionTest.roundup(1.0,0).should == 1.0
    FunctionTest.roundup(1.1,0).should == 2.0
    FunctionTest.roundup(1.5,0).should == 2.0
    FunctionTest.roundup(1.53,1).should == 1.6
    FunctionTest.roundup(-1.53,1).should == -1.6
  end

  it "should return an error when given inappropriate arguments" do
    FunctionTest.roundup("Asdasddf","asdfas").should == :value
  end
    
  it "should treat nil as zero" do
    FunctionTest.roundup(1.1,nil).should == 2.0
    FunctionTest.roundup(nil,1).should == 0
  end
  
  it "should return an error if an argument is an error" do
    FunctionTest.roundup(1,:error).should == :error
    FunctionTest.roundup(:error,1).should == :error
    FunctionTest.roundup(:error,:error).should == :error
  end
  
  it "should be in the list of functions that can be mapped to ruby" do
    MapFormulaeToRuby::FUNCTIONS[:'ROUNDUP'].should == 'roundup'
  end
  
end
