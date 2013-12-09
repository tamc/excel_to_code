require_relative '../../spec_helper.rb'

describe "ExcelFunctions: ROUND" do
  
  it "should round numbers correctly" do
    FunctionTest.round(1.1,0).should == 1.0
    FunctionTest.round(1.5,0).should == 2.0
    FunctionTest.round(1.56,1).should == 1.6
  end
  
  it "should work if arguments are given as strings, so long as the strings contain numbers" do
    FunctionTest.round('1.56','1').should == 1.6
  end
  
  it "should work if arguments given as booleans, with true = 1 and false = 0" do
    FunctionTest.round('1.56',true).should == 1.6
  end
    
  it "should treat nil as zero" do
    FunctionTest.round('1.56',nil).should == 2
    FunctionTest.round(nil,nil).should == 0
  end
    
  it "should return an error if an argument is an error" do
    FunctionTest.round(:error1,1).should == :error1
    FunctionTest.round(1,:error2).should == :error2
    FunctionTest.round(:error1,:error2).should == :error1
  end
  
  it "should be in the list of functions that can be mapped to ruby" do
    MapFormulaeToRuby::FUNCTIONS[:'ROUND'].should == 'round'
  end
  
end
