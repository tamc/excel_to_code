require_relative '../../spec_helper.rb'

describe "ExcelFunctions: EXP" do
  
  it "should return the constant e when given a power of 1" do
    FunctionTest.exp(1).should == Math::E
  end

  it "should return the the square of constant e when given a power of 2 (7.3890560989306495)" do
    FunctionTest.exp(2).should == Math::E ** 2
  end

  it "should return an error when given inappropriate arguments" do
    FunctionTest.exp("Asdasddf").should == :error
  end
    
  it "should treat nil as zero" do
    FunctionTest.exp(nil).should == 0
  end
  
  it "should return an error if an argument is an error" do
    FunctionTest.exp(:error).should == :error
  end
  
  it "should be in the list of functions that can be mapped to ruby" do
    MapFormulaeToRuby::FUNCTIONS['EXP'].should == 'exp'
  end
  
end
