require_relative '../../spec_helper.rb'

describe "ExcelFunctions: HYPERLINK" do
  
  it "should treat its arguments as strings, always" do
    FunctionTest.hyperlink("A").should == "<a href=\"A\">A</a>"
    FunctionTest.hyperlink(1).should == "<a href=\"1\">1</a>"
    FunctionTest.hyperlink(true).should == "<a href=\"TRUE\">TRUE</a>"
    FunctionTest.hyperlink(:div0).should == "<a href=\"#DIV/0!\">#DIV/0!</a>"
  end

  it "should accept an optional second argument to give it a friendly name" do
    FunctionTest.hyperlink("A","B").should == "<a href=\"A\">B</a>"
  end
  
  it "should be in the list of functions that can be mapped to ruby" do
    MapFormulaeToRuby::FUNCTIONS[:'HYPERLINK'].should == 'hyperlink'
  end
  
end
