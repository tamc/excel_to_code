require_relative '../../spec_helper.rb'

describe "ExcelFunctions: SCURVE" do
  
  it "scurve should return something when given appropriate arguments" do
    FunctionTest.scurve(2018,0,10,10).round.should == 0
    FunctionTest.scurve(2023,0,10,10).round.should == 5
    FunctionTest.scurve(2028,0,10,10).round.should == 10
  end

  it "scurve should cope with string arguments" do
    FunctionTest.scurve("2028","0","10","10").round.should == 10
  end

  it "halfscurve should return something when given appropriate arguments" do
    FunctionTest.halfscurve(2018,0,10,10).round.should == 0
    FunctionTest.halfscurve(2023,0,10,10).round.should == 5 
    FunctionTest.halfscurve(2028,0,10,10).round.should == 5
  end

  it "halfscurve should cope with string arguments" do
    FunctionTest.halfscurve("2028","0","10","10").round.should == 5
  end

  it "lcurve should return something when given appropriate arguments" do
    FunctionTest.lcurve(2018,0,10,10).round.should == 0
    FunctionTest.lcurve(2023,0,10,10).round.should == 5 
    FunctionTest.lcurve(2028,0,10,10).round.should == 10
  end

  it "lcurve should cope with string arguments" do
    FunctionTest.lcurve("2028","0","10","10").round.should == 10
  end

  it "curve should return something when given appropriate arguments" do
    FunctionTest.curve("hs", 2023,0,10,10).round.should == 5
    FunctionTest.curve("l",  2023,0,10,10).round.should == 5 
    FunctionTest.curve("s",  2023,0,10,10).round.should == 5
  end
  
end
