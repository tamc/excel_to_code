require_relative '../../spec_helper.rb'

describe "ExcelFunctions: INT" do
  
  it "should return 1 rounded to lowest integer (1.49)" do
    FunctionTest.int(1.49).should == 1
  end

  it "should return 3 rounded to lowest integer (3.0723)" do
    FunctionTest.int(3).should == 3
  end

  it "should return 7 rounded to lowest integer (7.9999999999)" do
    FunctionTest.int(7).should == 7
  end

  it "should return an error when given inappropriate arguments" do
    FunctionTest.int("Asdasddf").should == :error
  end
    
  it "should treat nil as zero" do
    FunctionTest.int(nil).should == 0
  end
  
  it "should return an error if an argument is an error" do
    FunctionTest.int(:error).should == :error
  end
  
  it "should be in the list of functions that can be mapped to ruby" do
    MapFormulaeToRuby::FUNCTIONS['INT'].should == 'int'
  end
  
end
