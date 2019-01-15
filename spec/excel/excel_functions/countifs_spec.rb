require_relative '../../spec_helper.rb'

describe "ExcelFunctions: COUNTIFS" do
  
  it "should only count values that meet all of the criteria" do
    FunctionTest.countifs([[10],[100],[nil]],10.0,[["pear"],["bear"],["apple"]],'Bear').should == 0.0
    FunctionTest.countifs([[10],[100],[nil]],10.0,[["pear"],["bear"],["apple"]],'Pear').should == 1.0
  end
  
  it "should work when single cells are given where ranges expected" do
    FunctionTest.countifs("CAR", "CAR", "FCV", "FCV").should == 1 
  end

  it "should match numbers with strings that contain numbers" do
    FunctionTest.countifs(10,"10.0").should == 1
    FunctionTest.countifs("10",10.0).should == 1
    FunctionTest.countifs([["CO2"],["CH4"],["N2O"],["CH4"],["N2O"],["CO2"]],"CO2",[["1A"],["1A"],["1A"],[4],[4],[5]],2).should == 0
  end
    
  it "should treat nil as an empty string when in the check_range, and a zero in the criteria" do
    FunctionTest.countifs(nil,20).should == 0
    FunctionTest.countifs(nil,"").should == 1 
    FunctionTest.countifs(nil,nil).should == 0
    FunctionTest.countifs(0,nil).should == 1
  end
  
  it "should deal with errors" do
    FunctionTest.countifs(:error2,20).should == 0
    FunctionTest.countifs(10,:error3).should == 0
    FunctionTest.countifs(:error1,:error1).should == 1
  end
  
  it "should be in the list of functions that can be mapped to ruby" do
    MapFormulaeToRuby::FUNCTIONS[:'COUNTIFS'].should == 'countifs'
  end
  
end
