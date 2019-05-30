require_relative '../../spec_helper.rb'

describe "ExcelFunctions: power(number,number)" do
  
  it "should return power of its arguments" do
    FunctionTest.power(2,0).should == 1
    FunctionTest.power(2,3).should == 8
    FunctionTest.power(4.0,0.5).should == 2.0
    FunctionTest.power(0.5,0.5).should ==0.7071067811865476 
  end

  it "should return NUM when trying to take roots of negative numbers" do
    FunctionTest.power(-4,2).should == 16
    FunctionTest.power(-4,0).should == 1
    FunctionTest.power(-4,0.5).should == :num
  end
      
  it "should treat nil as zero" do
    FunctionTest.power(1,nil).should == 1
    FunctionTest.power(nil,nil).should == 1
    FunctionTest.power(nil,1).should == 0
  end
    
  it "should work if numbers are given as strings" do
    FunctionTest.power("2","3").should == 8
  end
  
  # it "should be able to power arrays" do
  #   FunctionTest.power([[1,2],[3,4]],2).should == [[1,4],[9,16]]
  #   FunctionTest.power(2,[[1,2],[3,4]]).should == [[2,4],[8,16]]
  #   FunctionTest.power([[1,2],[3,4]],[[1,2],[3,4]]).should == [[1,4],[27,256]]
  # end
  
  it "should return an error if either argument is an error" do
    FunctionTest.power(:error,1).should == :error
    FunctionTest.power(1,:error).should == :error
    FunctionTest.power(:error1,:error2).should == :error1
  end

  it "should return num error if result is infinite" do
    FunctionTest.power(1e999, 1e999).should == :num
  end
  
  
end
