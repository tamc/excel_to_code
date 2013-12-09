require_relative '../../spec_helper.rb'

describe "ExcelFunctions: SUBTOTAL" do
  
  it "should calculate averages, counts, countas, sums depending on first argument" do
    FunctionTest.subtotal(1.0,1,"two",[[10],[100],[nil]]).should == 111.0/3.0 # Average
    FunctionTest.subtotal(2.0,1,"two",[[10],[100],[nil]]).should == 3 # count
    FunctionTest.subtotal(3.0,1,"two",[[10],[100],[nil]]).should == 4 # counta
    FunctionTest.subtotal(9.0,1,"two",[[10],[100],[nil]]).should == 111 # sum

    FunctionTest.subtotal(101.0,1,"two",[[10],[100],[nil]]).should == 111.0/3.0 # Average
    FunctionTest.subtotal(102.0,1,"two",[[10],[100],[nil]]).should == 3 # count
    FunctionTest.subtotal(103.0,1,"two",[[10],[100],[nil]]).should == 4 # counta
    FunctionTest.subtotal(109.0,1,"two",[[10],[100],[nil]]).should == 111 # sum    
  end

  it "the first argument can be a string, if the string contains a number" do
    FunctionTest.subtotal("1.0",1,"two",[[10],[100],[nil]]).should == 111.0/3.0 # Average
  end
  
  it "the first argument can be true, in whihc case it is interperted as asking for the average" do
    FunctionTest.subtotal(true,1,"two",[[10],[100],[nil]]).should == 111.0/3.0 # Average
  end

  it "returns :value if it requests an unsupported function number in the first argument" do
    FunctionTest.subtotal(0,1,"two",[[10],[100],[nil]]).should == :value
  end
    
  it "should return an error if the first argument is an error" do
    FunctionTest.subtotal(:error,1,"two",[[10],[100],[nil]]).should == :error
  end
  
  it "should be in the list of functions that can be mapped to ruby" do
    MapFormulaeToRuby::FUNCTIONS[:'SUBTOTAL'].should == 'subtotal'
  end
  
end
