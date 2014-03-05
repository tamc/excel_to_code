require_relative '../../spec_helper.rb'

describe "ExcelFunctions: SUMIF" do
  
  it "should only sum values in the area that meet the criteria" do
    FunctionTest.sumif(10,10.0).should == 10.0
    FunctionTest.sumif(10,10.0,20).should == 20
    FunctionTest.sumif([[10],[100],[nil]],10.0).should == 10.0
    FunctionTest.sumif([[10,100,nil]],10.0).should == 10.0
    FunctionTest.sumif([["pear"],["bear"],["apple"]],'Bear',[[10],[100],[nil]]).should == 100.0
  end
  
  it "should understand >0 type criteria" do
    FunctionTest.sumif([[10],[100],[nil]],">0").should == 110.0
    FunctionTest.sumif([[10],[100],[nil]],">10").should == 100.0
    FunctionTest.sumif([[10],[100],[nil]],"<100").should == 10.0
  end

  it "should match numbers with strings that contain numbers" do
    FunctionTest.sumif(10,"10.0").should == 10.0
  end
    
  it "should treat nil as an empty string when in the check_range, but not in the criteria" do
    FunctionTest.sumif(nil,0,20).should == 0
    FunctionTest.sumif(0,nil,100).should == 100
    FunctionTest.sumif(nil,nil, 100).should == 0
  end
  
  it "should deal with errors in arguments" do
    FunctionTest.sumif(:error1,10,20).should == 0
    FunctionTest.sumif(10,20,:error1).should == 0
    FunctionTest.sumif(20,20,:error1).should == :error1
    FunctionTest.sumif(1,:error2,20).should == 0
  end
  
  it "should be in the list of functions that can be mapped to ruby" do
    MapFormulaeToRuby::FUNCTIONS[:'SUMIF'].should == 'sumif'
  end
  
end
