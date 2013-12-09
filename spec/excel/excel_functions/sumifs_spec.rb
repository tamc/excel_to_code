require_relative '../../spec_helper.rb'

describe "ExcelFunctions: SUMIFS" do
  
  it "should only sum values that meet all of the criteria" do
    FunctionTest.sumifs([[10],[100],[nil]],[[10],[100],[nil]],10.0,[["pear"],["bear"],["apple"]],'Bear').should == 0.0
    FunctionTest.sumifs([[10],[100],[nil]],[[10],[100],[nil]],10.0,[["pear"],["bear"],["apple"]],'Pear').should == 10.0
  end
  
  it "should work when single cells are given where ranges expected" do
    FunctionTest.sumifs(0.143897265452564, "CAR", "CAR", "FCV", "FCV").should == 0.143897265452564
  end

  it "should match numbers with strings that contain numbers" do
    FunctionTest.sumifs(100,10,"10.0").should == 100
    FunctionTest.sumifs([[1],[2],[3],[4],[5],[5]],[["CO2"],["CH4"],["N2O"],["CH4"],["N2O"],["CO2"]],"CO2",[["1A"],["1A"],["1A"],[4],[4],[5]],2).should == 0
  end
    
  it "should treat nil as an empty string when in the check_range, but not in the criteria" do
    FunctionTest.sumifs(100,nil,20).should == 0
    FunctionTest.sumifs(100,nil,"").should == 100
    FunctionTest.sumifs(100,nil,nil).should == 0
  end
  
  it "should return an error if an argument is an error" do
    FunctionTest.sumifs(:error1,10,20).should == :error1
    FunctionTest.sumifs(1,:error2,20).should == :error2
    FunctionTest.sumifs(1,10,:error3).should == :error3
    FunctionTest.sumifs(:error1,:error2,:error3).should == :error1
  end
  
  it "should be in the list of functions that can be mapped to ruby" do
    MapFormulaeToRuby::FUNCTIONS[:'SUMIFS'].should == 'sumifs'
  end
  
end
