require_relative '../../spec_helper.rb'

describe "ExcelFunctions: EXACT" do
  
  it "Should perform a case sensitive equality check on strings" do
    FunctionTest.exact("word", "word").should == true
    FunctionTest.exact("Word", "word").should == false
    FunctionTest.exact("w ord", "word").should == false
  end

  it "Should convert other types to strings before comparing" do
    FunctionTest.exact("TRUE", true).should == true
    FunctionTest.exact("true", false).should == false

    FunctionTest.exact(1, "1").should == true
    FunctionTest.exact(1.1, "1").should == false
  end
    
  it "should treat nil as am empty string" do
    FunctionTest.exact(nil, "").should == true
    FunctionTest.exact(nil, "1").should == false
  end
  
  it "should return an error if an argument is an error" do
    FunctionTest.exact(:error, 1).should == :error
    FunctionTest.exact(1, :error).should == :error
  end
  
  it "should be in the list of functions that can be mapped to ruby" do
    MapFormulaeToRuby::FUNCTIONS[:'EXACT'].should == 'exact'
  end
  
end
