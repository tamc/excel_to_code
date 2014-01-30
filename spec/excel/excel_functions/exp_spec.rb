require_relative '../../spec_helper.rb'

describe "ExcelFunctions: EXP" do
  
  it "should return something when given appropriate arguments" do
    FunctionTest.exp(0).should == 1
    FunctionTest.exp(1).should == 2.718281828459045
    FunctionTest.exp("1").should == 2.718281828459045
    FunctionTest.exp(-1).should == 0.36787944117144233
    FunctionTest.exp(false).should == 1
    FunctionTest.exp(true).should == 2.718281828459045
  end

  it "should return an error when given inappropriate arguments" do
    FunctionTest.exp("Asdasddf").should == :value
  end
    
  it "should treat nil as zero" do
    FunctionTest.exp(nil).should == 1
  end
  
  it "should return an error if an argument is an error" do
    FunctionTest.exp(:error).should == :error
  end
  
  it "should be in the list of functions that can be mapped to ruby" do
    MapFormulaeToRuby::FUNCTIONS[:'EXP'].should == 'exp'
  end
  
end
