require_relative '../../spec_helper.rb'

describe "ExcelFunctions: PRODUCT" do
  
  it "should return the product of its arguments" do
    FunctionTest.product(1,2,3).should == 6
  end

  it "should return the product of its arguments, flattening arrays" do
    FunctionTest.product([[1],[2],[3]]).should == 6
  end
  
  it "should ignore any arguments that are not numbers" do
    FunctionTest.product(1,true,2,"Hello",3).should == 6
  end
  
  it "should skip nil" do
    FunctionTest.product(1,nil,2,nil,3).should == 6
    FunctionTest.product(nil).should == 0
    FunctionTest.product().should == 0
  end
  
  it "should return an error if any arguments are errors" do
    FunctionTest.sum([[1]],[[[[1,:name]]]]).should == :name
  end
  
  it "should be in the list of functions that can be mapped to ruby" do
    MapFormulaeToRuby::FUNCTIONS[:'PRODUCT'].should == 'product'
  end
  
end
