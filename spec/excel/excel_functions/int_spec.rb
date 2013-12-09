require_relative '../../spec_helper.rb'

describe "ExcelFunctions: INT" do
  
  it "should int numbers down to nearest integer correctly" do
    FunctionTest.int(8.9).should == 8.0
    FunctionTest.int(-8.9).should == -9.0
  end
  
  it "should work if arguments are given as strings, so long as the strings contain numbers" do
    FunctionTest.int('1.56').should == 1.0
  end
  
  it "should work if arguments given as booleans, with true = 1 and false = 0" do
    FunctionTest.int(true).should == 1.0
    FunctionTest.int(false).should == 0.0
  end
    
  it "should treat nil as zero" do
    FunctionTest.int(nil).should == 0
  end
    
  it "should return an error if an argument is an error" do
    FunctionTest.int(:error).should == :error
  end
  
  it "should be in the list of functions that can be mapped to ruby" do
    MapFormulaeToRuby::FUNCTIONS[:'INT'].should == 'int'
  end
  
end
