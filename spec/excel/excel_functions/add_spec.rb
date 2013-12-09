require_relative '../../spec_helper.rb'

describe "ExcelFunctions: add(number,number)" do
  
  it "should return sum of its arguments" do
    FunctionTest.add(1,1).should == 2
    FunctionTest.add(1.0,1.0).should == 2.0
  end
    
  it "should treat nil as zero" do
    FunctionTest.add(1,nil).should == 1
    FunctionTest.add(nil,nil).should == 0
    FunctionTest.add(nil,1).should == 1
  end
  
  it "should work if numbers are given as strings" do
    FunctionTest.add("1","1.0").should == 2.0
  end
  
  # it "should be able to add arrays" do
  #   FunctionTest.add([[1,2],[3,4]],1).should == [[2,3],[4,5]]
  #   FunctionTest.add(1,[[1,2],[3,4]]).should == [[2,3],[4,5]]
  #   FunctionTest.add([[1,2],[3,4]],[[1,2],[3,4]]).should == [[2,4],[6,8]]
  # end
  
  it "should return an error if either argument is an error" do
    FunctionTest.add(:error,1).should == :error
    FunctionTest.add(1,:error).should == :error
    FunctionTest.add(:error1,:error2).should == :error1
  end
  
  it "should be in the list of functions that can be mapped to ruby" do
    MapFormulaeToRuby::FUNCTIONS[:'+'].should == 'add'
  end
  
end
