require_relative '../../spec_helper.rb'

describe "ExcelFunctions: LN" do
  
  it "should return something when given appropriate arguments" do
    FunctionTest.ln(10).should == 2.302585092994046
    FunctionTest.ln(8).should == 2.0794415416798357
    FunctionTest.ln(0).should == :num
    FunctionTest.ln(-1).should == :num
  end

  it "should return an error when given inappropriate arguments" do
    FunctionTest.ln("Asdasddf").should == :value
  end
    
  it "should treat nil as zero" do
    FunctionTest.ln(nil).should == :num 
  end
  
  it "should return an error if an argument is an error" do
    FunctionTest.ln(:error).should == :error
  end
  
  it "should be in the list of functions that can be mapped to ruby" do
    MapFormulaeToRuby::FUNCTIONS[:'LN'].should == 'ln'
  end
  
end
