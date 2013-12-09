require_relative '../../spec_helper.rb'

describe "ExcelFunctions: MMULT" do
  
  it "should return something when given appropriate arguments" do
    FunctionTest.mmult([[1,2],[3,4]],[[4,3],[2,1]]).should == [[8,5],[20,13]]
    FunctionTest.mmult([[1,2]],[[3],[4]]).should == [[11]]
  end

  it "should return an error when given inappropriate arguments" do
    FunctionTest.mmult([[1,2],[3,'a']],[[4,3],[2,1]]).should == [[:value,:value],[:value,:value]]
  end
    
  it "should return an error if an argument is an error" do
    FunctionTest.mmult([[1,2],[3,:error]],[[4,3],[2,1]]).should == [[:value,:value],[:value,:value]]
    FunctionTest.mmult(:error,[[4,3],[2,1]]).should == :error
  end
  
  it "should return an error if ranges aren't the right size" do
    FunctionTest.mmult([[1,2],[3,4]],[[4,3]]).should == [[:value,:value],[:value,:value]]
  end

  it "should be in the list of functions that can be mapped to ruby" do
    MapFormulaeToRuby::FUNCTIONS[:'MMULT'].should == 'mmult'
  end
  
end
