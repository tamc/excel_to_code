require_relative '../../spec_helper.rb'

describe "ExcelFunctions: MID" do
  
  it "should return a substring of the first argument, starting at the character number of the second number and going on for the number of characters in the last argument" do
    FunctionTest.mid("ABCDEFGHIJK",1,1).should == "A"
    FunctionTest.mid("ABCDEFGHIJK",2,2).should == "BC"
    FunctionTest.mid("ABCDEFGHIJK",13,2).should == ""
    FunctionTest.mid("ABCDEFGHIJK",1,20).should == "ABCDEFGHIJK"
    FunctionTest.mid("ABCDEFGHIJK",1,nil).should == ""
    FunctionTest.mid("ABCDEFGHIJK",-1,10).should == :value
    FunctionTest.mid("ABCDEFGHIJK",1,-1).should == :value
  end

  it "should return an error when given inappropriate arguments" do
    FunctionTest.mid("ABCDEFGHIJK","Hi",1).should == :value
  end

  it "should convert its first argument to a string if it isn't" do
    FunctionTest.mid(12345,2,2).should == "23"
  end
    
  it "should return an error if an argument is an error" do
    FunctionTest.mid(:error, 1, 1).should == :error
    FunctionTest.mid("OK", :error, 1).should == :error
    FunctionTest.mid("OK", 1, :error).should == :error
  end
  
  it "should be in the list of functions that can be mapped to ruby" do
    MapFormulaeToRuby::FUNCTIONS[:'MID'].should == 'mid'
  end
  
end
