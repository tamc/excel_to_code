require_relative '../../spec_helper.rb'

describe "ExcelFunctions: PROJECT" do
  
  it "should return something when given appropriate arguments" do
    FunctionTest.project(1).should == 1
  end

  it "should return an error when given inappropriate arguments" do
    FunctionTest.project("Asdasddf").should == :value
  end
    
  it "should treat nil as zero" do
    FunctionTest.project(nil).should == 0
  end
  
  it "should return an error if an argument is an error" do
    FunctionTest.project(:error).should == :error
  end
  
  it "should be in the list of functions that can be mapped to ruby" do
    MapFormulaeToRuby::FUNCTIONS[:'PROJECT'].should == 'project'
  end
  
end
