require_relative '../../spec_helper.rb'

describe "ExcelFunctions: INDEX(array,row_number,[column_number])" do
  
  it "should return the value in the array at position row_number, column_number" do
    FunctionTest.index(10,1,1).should == 10
    FunctionTest.index([[10],["pear"]],2.0).should == "pear"
    FunctionTest.index([[10],[100],[nil]],2.0).should == 100.0
    FunctionTest.index([["pear"],["bear"],["apple"]],2.0).should == "bear"
    FunctionTest.index([[10,"pear"],[100,"bear"],[nil,"apple"]],1.0,2.0).should == "pear"
    FunctionTest.index([[10,"pear"],[100,"bear"],[nil,"apple"]],2.0,1.0).should == 100.0
    FunctionTest.index([[10,"pear"],[100,"bear"],[nil,"apple"]],3.0,1.0).should == 0.0
    FunctionTest.index([[10,"pear"],[100,"bear"],[nil,"apple"]],3.0,3.0).should == :ref
    FunctionTest.index([[10,"pear"],[100,"bear"],[nil,"apple"]],3.0).should == :ref
  end

  it "should return the whole row or column if given a zero row or column number" do 
    FunctionTest.index([[10,"pear"],[100,"bear"],[nil,"apple"]],1.0,0.0).should == [[10.0,"pear"]]
    FunctionTest.index([[10,"pear"],[100,"bear"],[nil,"apple"]],0.0,2.0).should == [["pear"],["bear"],["apple"]]
  end

  it "should return a :ref error when given arguments outside array range" do
    FunctionTest.index([[10],["pear"]],-1).should == :ref
    FunctionTest.index([[10],["pear"]],3).should == :ref
  end
    
  it "should treat nil as zero if given as a required row or column number" do
    FunctionTest.index([[10,"pear"],[100,"bear"],[nil,"apple"]],1.0,nil).should == [[10.0,"pear"]]
    FunctionTest.index([[10,"pear"],[100,"bear"],[nil,"apple"]],nil,2.0).should == [["pear"],["bear"],["apple"]]
  end
  
  it "should return an error if an argument is an error" do
    FunctionTest.index(:error,:error,:error).should == :error
  end
  
  it "should be in the list of functions that can be mapped to ruby" do
    MapFormulaeToRuby::FUNCTIONS[:'INDEX'].should == 'index'
  end
  
end
