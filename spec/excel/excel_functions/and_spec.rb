require_relative '../../spec_helper.rb'

describe "ExcelFunctions: AND (implemented as excel_and)" do
  
  it "should return the AND of the given arguments, which can be arrays" do
    FunctionTest.excel_and(true,true).should == true
    FunctionTest.excel_and(true,false).should == false
    FunctionTest.excel_and(false,true).should == false
    FunctionTest.excel_and(false,false).should == false
    FunctionTest.excel_and(true,true,true,true,true,false).should == false
    FunctionTest.excel_and([[true,true],[true,true]],[[false]]).should == false
  end
  
  it "should treat 0 values as false and 1 values as true" do
    FunctionTest.excel_and(1,1).should == true
    FunctionTest.excel_and(1,0).should == false
  end
  
  it "should ignore strings, other numbers and nils" do
    FunctionTest.excel_and(true,nil,"Hello",100,-3.1415).should == true
  end
  
  it "should return an error when there are no boolean or 1/0 arguments" do
    FunctionTest.excel_and("Asdasddf").should == :value
  end
      
  it "should return an error if an argument is an error" do
    FunctionTest.excel_and(:error).should == :error
  end
  
  it "should be in the list of functions that can be mapped to ruby" do
    MapFormulaeToRuby::FUNCTIONS[:'AND'].should == 'excel_and'
  end
  
end
