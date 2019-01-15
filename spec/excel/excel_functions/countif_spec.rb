require_relative '../../spec_helper.rb'

describe "ExcelFunctions: COUNTIF" do
  
  it "should only count values in the area that meet the criteria" do
    FunctionTest.countif(10,10.0).should == 1
    FunctionTest.countif([[10],[100],[nil]],10.0).should == 1
    FunctionTest.countif([[10,100,nil]],10.0).should == 1
    FunctionTest.countif([[["pear"],["bear"],["apple"]],'Bear',[[10],[100]]],[nil]).should == 0
  end
  
  it "should understand >0 type criteria" do
    FunctionTest.countif([[10],[100],[nil]],">0").should == 2
    FunctionTest.countif([[10],[100],[nil]],">10").should == 1
    FunctionTest.countif([[10],[100],[nil]],"<100").should == 1
    FunctionTest.countif([[10],[100],[nil]],"< 100").should == 1
  end

  it "should match numbers with strings that contain numbers" do
    FunctionTest.countif(10,"10.0").should == 1
  end
    
  it "should treat nil as an empty string if in the check range" do
    FunctionTest.countif([nil,""],100).should == 0
    FunctionTest.countif([nil,nil],100).should == 0
  end
  
  
  it "should be in the list of functions that can be mapped to ruby" do
    MapFormulaeToRuby::FUNCTIONS[:'COUNTIF'].should == 'countif'
  end
  
end
