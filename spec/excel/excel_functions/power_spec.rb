require_relative '../../spec_helper.rb'

describe "ExcelFunctions: power(number,number)" do
  
  it "should return sum of its arguments" do
    FunctionTest.power(2,3).should == 8
    FunctionTest.power(4.0,0.5).should == 2.0
  end
      
  it "should treat nil as zero" do
    FunctionTest.power(1,nil).should == 1
    FunctionTest.power(nil,nil).should == 1
    FunctionTest.power(nil,1).should == 0
  end
    
  it "should work if numbers are given as strings" do
    FunctionTest.power("2","3").should == 8
  end
  
  it "should be able to power arrays" do
    FunctionTest.power([[1,2],[3,4]],2).should == [[1,4],[9,16]]
    FunctionTest.power(2,[[1,2],[3,4]]).should == [[2,4],[8,16]]
    FunctionTest.power([[1,2],[3,4]],[[1,2],[3,4]]).should == [[1,4],[27,256]]
  end
  
  it "should return an error if either argument is an error" do
    FunctionTest.power(:error,1).should == :error
    FunctionTest.power(1,:error).should == :error
    FunctionTest.power(:error1,:error2).should == :error1
  end
  
end
