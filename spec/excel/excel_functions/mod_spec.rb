require_relative '../../spec_helper.rb'

describe "ExcelFunctions: MOD" do
  
  it "should return the remainder of a number" do
    FunctionTest.mod(10,3).should == 1.0
    FunctionTest.mod(10,5).should == 0.0
    FunctionTest.mod(1.1,1).should be_within(0.01).of(0.1)
  end
  
  it "should be possible for the the arguments to be strings, if they convert to a number" do
    FunctionTest.mod("3.5","2").should == 1.5
  end

  it "should treat nil as zero" do
    FunctionTest.mod(nil,10).should == 0
    FunctionTest.mod(10,nil).should == :div0
    FunctionTest.mod(nil,nil).should == :div0
  end
  
  it "should treat true as 1 and false as 0" do
    FunctionTest.mod(1.1,true).should  be_within(0.01).of(0.1)
    FunctionTest.mod(1.1,false).should == :div0
    FunctionTest.mod(false,10).should == 0
  end

  it "should return an error when given inappropriate arguments" do
    FunctionTest.mod("Asdasddf","adsfads").should == :value
  end
  
  it "should return an error if an argument is an error" do
    FunctionTest.mod(1,:error).should == :error
    FunctionTest.mod(:error,1).should == :error
    FunctionTest.mod(:error,:error).should == :error
  end
  
  it "should be in the list of functions that can be mapped to ruby" do
    MapFormulaeToRuby::FUNCTIONS[:'MOD'].should == 'mod'
  end
  
end
