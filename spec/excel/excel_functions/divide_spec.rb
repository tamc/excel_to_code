require_relative '../../spec_helper.rb'

describe "ExcelFunctions: divide(number,number)" do
  
  it "should return sum of its arguments" do
    FunctionTest.divide(4,2).should == 2
    FunctionTest.divide(1.0,0.5).should == 2.0
  end
  
  it "should return a :div0! if a divide by zero occurs" do
    FunctionTest.divide(1,0).should == :div0
  end
    
  it "should treat nil as zero" do
    FunctionTest.divide(1,nil).should == :div0
    FunctionTest.divide(nil,nil).should == :div0
    FunctionTest.divide(nil,1).should == 0
  end
    
  it "should work if numbers are given as strings" do
    FunctionTest.divide("4","2").should == 2
  end  
  
  # it "should be able to divide arrays" do
  #   FunctionTest.divide([[2,4],[6,8]],2).should == [[1,2],[3,4]]
  #   FunctionTest.divide(1,[[1,2],[3,4]]).should == [[1,0.5],[1/3.0,0.25]]
  #   FunctionTest.divide([[1,2],[3,4]],[[1,2],[3,4]]).should == [[1,1],[1,1]]
  # end
  
  it "should return an error if either argument is an error" do
    FunctionTest.divide(:error,1).should == :error
    FunctionTest.divide(1,:error).should == :error
    FunctionTest.divide(:error1,:error2).should == :error1
  end
  
end
