require_relative '../../spec_helper.rb'

describe "ExcelFunctions: LOG" do
  
  it "should return something when given appropriate arguments" do
    FunctionTest.log(10).should == 1
    FunctionTest.log(10,10).should == 1
    FunctionTest.log(8,2).should == 3
    FunctionTest.log(0,2).should == :num
  end

  it "should return an error when given inappropriate arguments" do
    FunctionTest.log("Asdasddf").should == :value
  end
    
  it "should treat nil as zero" do
    FunctionTest.log(nil).should == :num 
  end
  
  it "should return an error if an argument is an error" do
    FunctionTest.log(:error).should == :error
  end
  
  it "should be in the list of functions that can be mapped to ruby" do
    MapFormulaeToRuby::FUNCTIONS[:'LOG'].should == 'log'
  end
  
end
