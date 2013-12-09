require_relative '../../spec_helper.rb'

describe "ExcelFunctions: CHOOSE(index,arg1,arg2,arg3)" do
  
  it "should return index-th numbered argument, where index is the first argument" do
    FunctionTest.choose(1,10,20,30).should == 10
    FunctionTest.choose(2,10,20,30).should == 20
    FunctionTest.choose(1,[10,20,30],40,50).should == [10,20,30]
  end

  it "should return an error when given an index that is outside the argument range or is nil" do
    FunctionTest.choose(0,10,20,30).should == :value
    FunctionTest.choose(-1,10,20,30).should == :value
    FunctionTest.choose(4,10,20,30).should == :value
    FunctionTest.choose(nil,10,20,30).should == :value
  end
    
  it "should treat nil as zero if otherwise given in an argument" do
    FunctionTest.choose(2,10,nil,30).should == 0
  end
  
  it "should return an error if an argument is an error" do
    FunctionTest.choose(:error).should == :error
  end
  
  it "should be in the list of functions that can be mapped to ruby" do
    MapFormulaeToRuby::FUNCTIONS[:'CHOOSE'].should == 'choose'
  end
  
end
