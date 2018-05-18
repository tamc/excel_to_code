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
    FunctionTest.sumifs(100,"10",10.0).should == 100
    FunctionTest.sumifs([[1],[2],[3],[4],[5],[5]],[["CO2"],["CH4"],["N2O"],["CH4"],["N2O"],["CO2"]],"CO2",[["1A"],["1A"],["1A"],[4],[4],[5]],2).should == 0
  end

  it "should treat nil as an empty string when in the check_range, and a zero in the criteria" do
    FunctionTest.sumifs(100,nil,20).should == 0
    FunctionTest.sumifs(100,nil,"").should == 100
    FunctionTest.sumifs(100,nil,nil).should == 0
    FunctionTest.sumifs(100,0,nil).should == 100
  end

  it "should cope with * in string" do
    FunctionTest.sumifs([[1],[3],[5]],[["pear"],["bear"],["apple"]],'*ear').should == 4
    FunctionTest.sumifs([[1],[3],[5]],[["pear"],["bear"],["apple"]],'be*').should == 3
    FunctionTest.sumifs([[1],[3],[5]],[["pear"],["bear"],["apple"]],'*pp*').should == 5
    FunctionTest.sumifs([[1],[3],[5]],[["pear"],["bear"],["apple"]],'*xx*').should == 0
  end

  it "should deal with errors" do
    FunctionTest.sumifs(:error1,10,20).should == 0
    FunctionTest.sumifs(:error1,20,20).should == :error1
    FunctionTest.sumifs(1,:error2,20).should == 0
    FunctionTest.sumifs(1,10,:error3).should == 0
    FunctionTest.sumifs(1,:error1,:error1).should == 1
  end

  it "should be in the list of functions that can be mapped to ruby" do
    MapFormulaeToRuby::FUNCTIONS[:'SUMIFS'].should == 'sumifs'
  end

end
