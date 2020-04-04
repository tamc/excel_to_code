require_relative '../../spec_helper.rb'

describe "ExcelFunctions: INTERPOLATE" do
  
  it "Should take a range with a size of four and an integer index and return the value indexed" do
    FunctionTest.interpolate([[9,2,8,3]],1).should == 9
    FunctionTest.interpolate([[9,2,8,3]],2).should == 2
    FunctionTest.interpolate([[9,2,8,3]],3).should == 8
    FunctionTest.interpolate([[9,2,8,3]],4).should == 3
  end

  it "Should work with single columns and single rows" do
    FunctionTest.interpolate([[9,2,8,3]],2).should == 2
    FunctionTest.interpolate([[9],[2],[8],[3]],2).should == 2
  end

  it "Should return an error if the index is out of bounds" do
    FunctionTest.interpolate([[9,2,8,3]],0).should == :value
    FunctionTest.interpolate([[9,2,8,3]],0.9).should == :value
    FunctionTest.interpolate([[9,2,8,3]],-1).should == :value
    FunctionTest.interpolate([[9,2,8,3]],5.1).should == :value
    FunctionTest.interpolate([[9,2,8,3]],6).should == :value
  end

  it "Should take a range with a size of four and a non-integer index and return a value interpolated between the two nearest values" do
    FunctionTest.interpolate([[9,2,8,3]],1.5).should be_within(0.01).of(5.5)
    FunctionTest.interpolate([[9,2,8,3]],2.1).should be_within(0.01).of(2.6)
    FunctionTest.interpolate([[9,2,8,3]],3.9).should be_within(0.01).of(3.5)
  end
  
  it "should return an error if an argument is an error" do
    FunctionTest.interpolate(:error, 1).should == :error
    FunctionTest.interpolate([[1]], :error).should == :error
  end
  
  it "should be in the list of functions that can be mapped to ruby" do
    MapFormulaeToRuby::FUNCTIONS[:'INTERPOLATE'].should == 'interpolate'
  end
  
end
