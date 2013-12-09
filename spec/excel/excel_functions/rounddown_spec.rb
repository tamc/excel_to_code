require_relative '../../spec_helper.rb'

describe "ExcelFunctions: ROUNDDOWN(number,decimal places)" do
  
  it "should round numbers down" do
    FunctionTest.rounddown(1.0,0).should == 1.0
    FunctionTest.rounddown(1.1,0).should == 1.0
    FunctionTest.rounddown(1.5,0).should == 1.0
    FunctionTest.rounddown(1.53,1).should == 1.5
    FunctionTest.rounddown(-1.53,1).should == -1.5
  end

  it "should return an error when given inappropriate arguments" do
    FunctionTest.rounddown("Asdasddf","asdfas").should == :value
  end
    
  it "should treat nil as zero" do
    FunctionTest.rounddown(1.1,nil).should == 1.0
    FunctionTest.rounddown(nil,1).should == 0
  end
  
  it "should return an error if an argument is an error" do
    FunctionTest.rounddown(1,:error).should == :error
    FunctionTest.rounddown(:error,1).should == :error
    FunctionTest.rounddown(:error,:error).should == :error
  end
  
  it "should be in the list of functions that can be mapped to ruby" do
    MapFormulaeToRuby::FUNCTIONS[:'ROUNDDOWN'].should == 'rounddown'
  end
  
end
