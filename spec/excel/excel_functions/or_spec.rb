require_relative '../../spec_helper.rb'

describe "ExcelFunctions: OR" do
  
  it "should return the OR of the given arguments, which can be arrays" do
    FunctionTest.excel_or(true,true).should == true
    FunctionTest.excel_or(true,false).should == true
    FunctionTest.excel_or(false,true).should == true
    FunctionTest.excel_or(false,false).should == false
    FunctionTest.excel_or(true,true,true,true,true,false).should == true
    FunctionTest.excel_or(false,false,false,false,false,false).should == false
    FunctionTest.excel_or([[true,true],[true,true]],[[false]]).should == true
  end
  
  it "should treat 0 values as false or 1 values as true" do
    FunctionTest.excel_or(1,1).should == true
    FunctionTest.excel_or(1,0).should == true
    FunctionTest.excel_or(0,0).should == false
  end
  
  it "should ignore strings, other numbers or nils" do
    FunctionTest.excel_or(true,nil,"Hello",100,-3.1415).should == true
    FunctionTest.excel_or(false,nil,"Hello",100,-3.1415).should == false
  end
  
  it "should return an error when there are no boolean or 1/0 arguments" do
    FunctionTest.excel_or("Asdasddf").should == :value
  end
      
  it "should return an error if an argument is an error" do
    FunctionTest.excel_or(:error).should == :error
  end
  
  it "should be in the list of functions that can be mapped to ruby" do
    MapFormulaeToRuby::FUNCTIONS[:'OR'].should == 'excel_or'
  end
  
end
