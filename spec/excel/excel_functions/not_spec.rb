require_relative '../../spec_helper.rb'

describe "ExcelFunctions: NOT (implemented as excel_not)" do
  
  it "should return the NOT of the given argument " do
    FunctionTest.excel_not(true).should == false
    FunctionTest.excel_not(false).should == true
  end
  
  it "should treat 0 values as false and other numbers as true" do
    FunctionTest.excel_not(10).should == false
    FunctionTest.excel_not(1).should == false
    FunctionTest.excel_not(0).should == true
  end
  
  it "should treat blanks as false" do
    FunctionTest.excel_not(nil).should == true
  end
  
  it "should return an error when given a string" do
    FunctionTest.excel_not("Asdasddf").should == :value
  end
      
  it "should return an error if an argument is an error" do
    FunctionTest.excel_not(:error).should == :error
  end

  it "should return an error if an argument is a range" do
    FunctionTest.excel_not([true,false]).should == :value
  end
  
  it "should be in the list of functions that can be mapped to ruby" do
    MapFormulaeToRuby::FUNCTIONS[:'NOT'].should == 'excel_not'
  end
  
end
