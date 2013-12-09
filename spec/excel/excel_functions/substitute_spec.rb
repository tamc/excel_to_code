require_relative '../../spec_helper.rb'

describe "ExcelFunctions: SUBSTITUTE" do
  
  it "it should replace any occurences of its second argument in its first argument with its third argument. If its fourth argument is emptly, all arguments replaced, otherwise the ioccurence given by the number in its fourth argument is replaced. Case matters. Non strings are converted into strings." do
    FunctionTest.substitute(nil,nil,nil).should == ""
    FunctionTest.substitute("hello",nil,"world").should == "hello"
    FunctionTest.substitute("hello",'ll',"world").should == "heworldo"
    FunctionTest.substitute("hello",'LL',"world").should == "hello"
    FunctionTest.substitute("heTRUEo",true,"world").should == "heworldo"
    FunctionTest.substitute("he3o",3,"world").should == "heworldo"
    FunctionTest.substitute("ABABABAB","AB","CD").should == "CDCDCDCD"
    FunctionTest.substitute("ABABABAB","AB","CD", 1).should == "CDABABAB"
    FunctionTest.substitute("ABABABAB","AB","CD", 2).should == "ABCDABAB"
    FunctionTest.substitute("ABABABAB","AB","CD", 20).should == "ABABABAB"
  end

  it "should return an error when given inappropriate arguments" do
    FunctionTest.substitute("ABABABAB","AB","CD", 0).should == :value
    FunctionTest.substitute("ABABABAB","AB","CD", "Hi").should == :value
  end
    
  it "should return an error if an argument is an error" do
    FunctionTest.substitute(:error, 'll', 'world').should == :error
    FunctionTest.substitute("hello", :error, 'world').should == :error
    FunctionTest.substitute("hello", 'll', :error).should == :error
    FunctionTest.substitute("he3o",3,"world", :error).should == :error
  end
  
  it "should be in the list of functions that can be mapped to ruby" do
    MapFormulaeToRuby::FUNCTIONS[:'SUBSTITUTE'].should == 'substitute'
  end
  
end
