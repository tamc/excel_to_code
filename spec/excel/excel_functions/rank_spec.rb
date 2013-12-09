require_relative '../../spec_helper.rb'

describe "ExcelFunctions: RANK" do
  
  it "should return the rank of the first argument in the list given as the second argument assuming the order specified in the third argument" do
    FunctionTest.rank(2,[[3,2],[1,8]]).should == 3
    FunctionTest.rank(2,[[3,2],[1,8]],0).should == 3
    FunctionTest.rank(2,[[3,2],[1,8]],1).should == 2
    # Should ignore nonumeric values, including numbers in strings and booleans
    FunctionTest.rank(2,[[3,"Hi!", 2],[1,8]],1).should == 2
    FunctionTest.rank(2,[[3,false, 2],[1,8]],1).should == 2
    # Should give duplicate number the same rank
    FunctionTest.rank(2,[[3, 2, 2],[1,8]],1).should == 2
    FunctionTest.rank(2,[[3, 2, 1],[1,8]],1).should == 3
    # Should return #N/A if number doesn't appear in list
    FunctionTest.rank(1.5,[[3,2],[1,8]],1).should == :na
  end

  it "should return an error when given the wrong sort of argument" do
    FunctionTest.rank("2",[[3,2],[1,8]]).should == 3
    FunctionTest.rank("Hi",[[3,2],[1,8]]).should == :value
    FunctionTest.rank(2,2).should == 1
  end
    
  it "should return an error if an argument is an error or list passed has an error" do
    FunctionTest.rank(1, :error).should == :error
    FunctionTest.rank(:error, [[1,2]]).should == :error
    FunctionTest.rank(2,[[3,2],[:na,8]],1).should == :na
  end
  
  it "should be in the list of functions that can be mapped to ruby" do
    MapFormulaeToRuby::FUNCTIONS[:'RANK'].should == 'rank'
  end
  
end
