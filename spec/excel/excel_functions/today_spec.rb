require_relative '../../spec_helper.rb'

describe "ExcelFunctions: TODAY" do
  
  it "should return something when given appropriate arguments" do
    FunctionTest.today(1).should == 1
  end

  it "should return an error when given inappropriate arguments" do
    FunctionTest.today("Asdasddf").should == :value
  end
    
  it "should treat nil as zero" do
    FunctionTest.today(nil).should == 0
  end
  
  it "should return an error if an argument is an error" do
    FunctionTest.today(:error).should == :error
  end
  
  it "should be in the list of functions that can be mapped to ruby" do
    MapFormulaeToRuby::FUNCTIONS['TODAY'].should == 'today'
  end
  
end
