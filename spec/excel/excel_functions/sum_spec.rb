require_relative '../../spec_helper.rb'

describe "ExcelFunctions: SUM()" do
  
  it "should return the sum of its arguments" do
    FunctionTest.sum(1,2,3).should == 6
  end

  it "should return the sum of its arguments, flattening arrays" do
    FunctionTest.sum([[1],[2],[3]]).should == 6
  end
  
  it "should ignore any arguments that are not numbers" do
    FunctionTest.sum(1,true,2,"Hello",3).should == 6
  end
  
  it "should treat nil as zero" do
    FunctionTest.sum(1,nil,2,nil,3).should == 6
    FunctionTest.sum(nil).should == 0
    FunctionTest.sum().should == 0
  end
  
  it "should return an error if any arguments are errors" do
    FunctionTest.sum([[1]],[[[[1,:name]]]]).should == :name
  end
  
end
