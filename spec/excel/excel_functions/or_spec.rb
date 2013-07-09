require_relative '../../spec_helper.rb'

describe "ExcelFunctions: OR" do
  
  it "should return something when given appropriate arguments" do
    FunctionTest.or(1).should == 1
  end

  it "should return an error when given inappropriate arguments" do
    FunctionTest.or("Asdasddf").should == :value
  end
    
  it "should treat nil as zero" do
    FunctionTest.or(nil).should == 0
  end
  
  it "should return an error if an argument is an error" do
    FunctionTest.or(:error).should == :error
  end
  
  it "should be in the list of functions that can be mapped to ruby" do
    MapFormulaeToRuby::FUNCTIONS['OR'].should == 'or'
  end
  
end
