require_relative '../../spec_helper.rb'

describe "negative(function) should be equivalent to -function" do
  
  it "should return the negative of its arguments" do
    FunctionTest.negative(1).should == -1
    FunctionTest.negative(-1).should == 1
  end
  
  it "should treat strings that only contain numbers as numbers" do
    FunctionTest.negative("10").should == -10
    FunctionTest.negative("-1.3").should == 1.3
  end

  it "should return an error when given inappropriate arguments" do
    FunctionTest.negative("Asdasddf").should == :value
  end
    
  it "should treat nil as zero" do
    FunctionTest.negative(nil).should == 0
  end
  
  it "should return an error if an argument is an error" do
    FunctionTest.negative(:error).should == :error
  end
  
end
