require_relative '../../spec_helper.rb'

describe "ExcelFunctions: TRIM" do
  
  it "should return text with all leading and trailing space removed and only single spaces within" do
    FunctionTest.trim(" This is a long        space   ").should == "This is a long space"
    FunctionTest.trim("Thisisok").should == "Thisisok"
  end

  it "should return any non string arguments as they are" do
    FunctionTest.trim(12).should == 12
    FunctionTest.trim(false).should == false
  end
  
  it "should be in the list of functions that can be mapped to ruby" do
    MapFormulaeToRuby::FUNCTIONS[:'TRIM'].should == 'trim'
  end
  
end
