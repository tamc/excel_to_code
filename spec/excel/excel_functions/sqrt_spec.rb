require_relative '../../spec_helper.rb'

describe "ExcelFunctions: SQRT" do
  
  it "should return something when given appropriate arguments" do
    FunctionTest.sqrt(1).should == 1
    FunctionTest.sqrt(4).should == 2
    FunctionTest.sqrt(1.6).should == 1.2649110640673518
  end

  it "should return an error when given inappropriate arguments" do
    FunctionTest.sqrt("Asdasddf").should == :value
    FunctionTest.sqrt(-1).should == :num
  end
    
  it "should treat nil as zero" do
    FunctionTest.sqrt(nil).should == 0
  end
  
  it "should return an error if an argument is an error" do
    FunctionTest.sqrt(:error).should == :error
  end
  
  it "should be in the list of functions that can be mapped to ruby" do
    MapFormulaeToRuby::FUNCTIONS[:'SQRT'].should == 'sqrt'
  end
  
end
