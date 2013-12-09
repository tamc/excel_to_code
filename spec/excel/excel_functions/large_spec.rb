require_relative '../../spec_helper.rb'

describe "ExcelFunctions: LARGE(range, k)" do
  
  it "should return the kth largest value in the range number" do
    FunctionTest.large([[10],[100],[500],[nil]],1).should == 500
    FunctionTest.large([[10],[100],[500],[nil]],2).should == 100
    FunctionTest.large([[10],[100],[500],[nil]],3).should == 10
    FunctionTest.large([[10],[100],[500],[nil]],4).should == :num
  end

  it "should work when passed a single cell instead of a range" do
    FunctionTest.large(500,1).should == 500
  end

  it "should return errors" do
    FunctionTest.large([[10],[100],[500],[:value]],4).should == :value
    FunctionTest.large([[10],[100],[500],[nil]],:div0).should == :div0
  end
  
  it "should be in the list of functions that can be mapped to ruby" do
    MapFormulaeToRuby::FUNCTIONS[:'LARGE'].should == 'large'
  end
  
end
