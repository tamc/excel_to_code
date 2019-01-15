require_relative '../../spec_helper.rb'

describe "ExcelFunctions: REPLACE" do
  
  it "should replace some characters in a string" do
    FunctionTest.replace("1234",1,1,'@').should == "@234"
    FunctionTest.replace("1234",2,2,'@').should == "1@4"
    FunctionTest.replace("1234",10,2,'@').should == "1234@"
  end

  it "should return an error when given inappropriate arguments" do
    FunctionTest.replace(1234,2,2,'@').should == "1@4"
  end
    
  it "should treat nil as empty string" do
    FunctionTest.replace(nil,10,10,"@").should == "@" 
  end
  
  it "should return an error if an argument is an error" do
    FunctionTest.replace(:error, 2, 2, '@').should == :error
  end
  
  it "should be in the list of functions that can be mapped to ruby" do
    MapFormulaeToRuby::FUNCTIONS[:'REPLACE'].should == 'replace'
  end
  
end
