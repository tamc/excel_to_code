require_relative '../../spec_helper.rb'

describe "ExcelFunctions: MONTH" do
  
  it "should return 3 when date is 03/26/2013 (serial: 41359)" do
    FunctionTest.month(41359).should == 3
  end

  it "should return an error when given inappropriate arguments" do
    FunctionTest.month("Asdasddf").should == :error
  end
    
  it "should treat nil as zero" do
    FunctionTest.month(nil).should == 0
  end
  
  it "should return an error if an argument is an error" do
    FunctionTest.month(:error).should == :error
  end
  
  it "should be in the list of functions that can be mapped to ruby" do
    MapFormulaeToRuby::FUNCTIONS['MONTH'].should == 'month'
  end
  
end
