require_relative '../../spec_helper.rb'

describe "ExcelFunctions: NPV" do
  
  it "should return something when given appropriate arguments" do
    FunctionTest.npv(0.1, 110).should == 99.99999999999999
    FunctionTest.npv(0.1, 110, 121).should == 199.99999999999997
    FunctionTest.npv(0.1, [[110], [121]]).should == 199.99999999999997
  end

  it "should return an error when given inappropriate arguments" do
    FunctionTest.npv(-1, 110).should == :div0
    FunctionTest.npv("Asdasddf").should == :value
  end
    
  it "should treat nil as zero" do
    FunctionTest.npv(nil,100).should == 100
    FunctionTest.npv(0.1,nil).should == 0
  end
  
  it "should return an error if an argument is an error" do
    FunctionTest.npv(:error, 100).should == :error
    FunctionTest.npv(0.1, :error).should == :error
  end
  
  it "should be in the list of functions that can be mapped to ruby" do
    MapFormulaeToRuby::FUNCTIONS[:'NPV'].should == 'npv'
  end
  
end
