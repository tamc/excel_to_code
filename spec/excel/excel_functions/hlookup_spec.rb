require_relative '../../spec_helper.rb'

describe "ExcelFunctions: HLOOKUP" do
  
  it "should match the first argument against the first row of the table in the second argument, returning the value in the row specified by the third argument" do
    test = [[1,2,3],['a','b','c']]
    FunctionTest.hlookup(2.0,test,2).should == 'b'
    FunctionTest.hlookup(1.5,test,2).should == 'a'
    FunctionTest.hlookup(0.5,test,2).should == :na
    FunctionTest.hlookup(10,test,2).should == 'c'
    FunctionTest.hlookup(2.6,test,2).should == 'b'
    FunctionTest.hlookup(2.6,test,2,true).should == 'b'
    FunctionTest.hlookup(2.6,test,2,false).should == :na

    test = [["hello",2,3],['a','b','c']]
    FunctionTest.hlookup("HELLO", test, 2, false).should == 'a'
    FunctionTest.hlookup("HELMP", test, 2, true).should == 'a'
  end
    
  it "nil should not match with anything" do
    FunctionTest.hlookup(nil,[[nil,'a'],[2,'b'],[3,'c']],2).should == :na
  end
  
  it "should return an error if an argument is an error" do
    FunctionTest.hlookup(:error,[[1,'a'],[2,'b'],[3,'c']],2,false).should == :error
    FunctionTest.hlookup(2.0,:error,2,false).should == :error
    FunctionTest.hlookup(2.0,[[1,'a'],[2,'b'],[3,'c']],:error,false).should == :error
    FunctionTest.hlookup(2.0,[[1,'a'],[2,'b'],[3,'c']],2,:error).should == :error
    FunctionTest.hlookup(:error,:error,:error,:error).should == :error
  end
  
  it "should be in the list of functions that can be mapped to ruby" do
    MapFormulaeToRuby::FUNCTIONS[:'HLOOKUP'].should == 'hlookup'
  end
  
end
