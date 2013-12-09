require_relative '../../spec_helper.rb'

describe "ExcelFunctions: TEXT" do
  
  it "should turn a number into a percentage when the second argument is 0%" do 
    FunctionTest.text(1, "0%").should == "100%"
    FunctionTest.text(0.196, "0%").should == "20%"
    FunctionTest.text("0.196", "0%").should == "20%"
  end

  it "should pass non-numbers through unchanged" do
    FunctionTest.text("Asdasddf","0%").should == "Asdasddf"
  end
    
  it "should treat nil as zero if given as number" do
    FunctionTest.text(nil,"0%").should == "0%"
  end

  it "should treat nil as empty string if given as format" do
    FunctionTest.text(1, nil).should == ""
  end
  
  it "should pass on error if provided as an argument" do
    FunctionTest.text(:error,"0%").should == :error
    FunctionTest.text(100,:error).should == :error
  end

  it "should be in the list of functions that can be mapped to ruby" do
    MapFormulaeToRuby::FUNCTIONS[:'TEXT'].should == 'text'
  end
  
end
