require_relative '../../spec_helper.rb'

describe "ExcelFunctions: VLOOKUP" do
  
  it "should match the first argument against the first column of the table in the second argument, returning the value in the column specified by the third argument" do
    FunctionTest.vlookup(2.0,[[1,'a'],[2,'b'],[3,'c']],2).should == 'b'
    FunctionTest.vlookup(1.5,[[1,'a'],[2,'b'],[3,'c']],2).should == 'a'
    FunctionTest.vlookup(0.5,[[1,'a'],[2,'b'],[3,'c']],2).should == :na
    FunctionTest.vlookup(10,[[1,'a'],[2,'b'],[3,'c']],2).should == 'c'
    FunctionTest.vlookup(2.6,[[1,'a'],[2,'b'],[3,'c']],2).should == 'b'
    FunctionTest.vlookup(2.6,[[1,'a'],[2,'b'],[3,'c']],2,true).should == 'b'
    FunctionTest.vlookup(2.6,[[1,'a'],[2,'b'],[3,'c']],2,false).should == :na
    FunctionTest.vlookup("HELLO",[['hello','a'],[2,'b'],[3,'c']],2,false).should == 'a'
    FunctionTest.vlookup("HELMP",[['hello','a'],[2,'b'],[3,'c']],2,true).should == 'a'
  end
    
  it "nil should not match with anything" do
    FunctionTest.vlookup(nil,[[nil,'a'],[2,'b'],[3,'c']],2).should == :na
  end
  
  it "should return an error if an argument is an error" do
    FunctionTest.vlookup(:error,[[1,'a'],[2,'b'],[3,'c']],2,false).should == :error
    FunctionTest.vlookup(2.0,:error,2,false).should == :error
    FunctionTest.vlookup(2.0,[[1,'a'],[2,'b'],[3,'c']],:error,false).should == :error
    FunctionTest.vlookup(2.0,[[1,'a'],[2,'b'],[3,'c']],2,:error).should == :error
    FunctionTest.vlookup(:error,:error,:error,:error).should == :error
  end
  
  it "should be in the list of functions that can be mapped to ruby" do
    MapFormulaeToRuby::FUNCTIONS[:'VLOOKUP'].should == 'vlookup'
  end
  
end
