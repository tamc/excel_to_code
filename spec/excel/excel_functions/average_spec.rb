require_relative '../../spec_helper.rb'

describe "ExcelFunctions: AVERAGE()" do
  
  it "should return the average of its arguments" do
    FunctionTest.average(1,2,3).should == 2
    FunctionTest.average(1,2).should == 1.5
  end

  it "should return the average of its arguments, flattening arrays" do
    FunctionTest.average([[1],[2],[3]]).should == 2
  end
  
  it "should ignore any arguments that are not numbers" do
    FunctionTest.average(1,true,2,"Hello",3).should == 2
  end
  
  it "should treat nil as zero" do
    FunctionTest.average(1,nil,2,nil,3).should == 2
    FunctionTest.average(nil).should == :div0
  end
  
  it "should return an error if any arguments are errors" do
    FunctionTest.average([[1]],[[[[1,:name]]]]).should == :name
  end
  
end
