require_relative '../../spec_helper.rb'

require 'date'

describe "ExcelFunctions: DATE" do
  
  it "should return a valid date when given a valid year, month, day" do
    FunctionTest.date(1900,1,1).should === Date.new(1900,1,1)
  end

  it "should return an error when given inappropriate arguments" do
    FunctionTest.date("Asdasddf").should == :value
  end
    
  it "should treat nil as zero" do
    FunctionTest.date(nil).should == 0
  end
  
  it "should return an error if an argument is an error" do
    FunctionTest.date(:error).should == :error
  end
  
  it "should be in the list of functions that can be mapped to ruby" do
    MapFormulaeToRuby::FUNCTIONS[:'DATE'].should == 'date'
  end
  
end
