require_relative '../../spec_helper.rb'

describe "ExcelFunctions: UNICODE" do
  
  it "should return the codepoint of the first character in a string" do
    FunctionTest.unicode("A").should == 65
    FunctionTest.unicode("B").should == 66
    FunctionTest.unicode("ㇴ").should == 12788
  end

  it "should convert numbers and booleans to strings" do
    FunctionTest.unicode(1).should == 49
    FunctionTest.unicode(TRUE).should == 84
  end

  it "should return an error when given inappropriate arguments" do
    FunctionTest.unicode(nil).should == :value
  end
    
  it "should return an error if an argument is an error" do
    FunctionTest.unicode(:error).should == :error
  end
  
  it "should be in the list of functions that can be mapped to ruby" do
    MapFormulaeToRuby::FUNCTIONS[:'UNICODE'].should == 'unicode'
  end
  
end
