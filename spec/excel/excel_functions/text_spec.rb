require_relative '../../spec_helper.rb'

describe "ExcelFunctions: TEXT" do
  
  it "should turn a number into a percentage when the second argument is 0%" do 
    FunctionTest.text(1, "0%").should == "100%"
    FunctionTest.text(0.196, "0%").should == "20%"
    FunctionTest.text("0.196", "0%").should == "20%"
    FunctionTest.text("0.196", "0.0%").should == "19.6%"
  end

  it "should turn a number into a float with the appropriate number of decimal places if 0.0" do
    FunctionTest.text(1.184, "0").should == "1"
    FunctionTest.text(1.184, 0).should == "1"
    FunctionTest.text(1.184, "0.0").should == "1.2"
    FunctionTest.text(1.134, "0.0").should == "1.1"
    FunctionTest.text(1.134, "0.00").should == "1.13"
    FunctionTest.text(1.134, "0.000").should == "1.134"
  end

  it "should turn a number into a float commas separating thousands if #,##" do
    FunctionTest.text(123456789.123456, "#,##").should == "123,456,789"
    FunctionTest.text(123456789.123456, "#,##0").should == "123,456,789"
    FunctionTest.text(123456789.123456, "#,##0.0").should == "123,456,789.1"
  end

  it "should turn a number into an integer with extra zeros if needed if 0000" do
    # FIXME: Excel rounds up, programs see to round to even
    FunctionTest.text(12.51, "0000").should == "0013"
    FunctionTest.text(12500.51, "0000").should == "12501"
  end

  it "Should cope with international variants" do
    FunctionTest.text("0.196", "0,0%").should == "19.6%"
    FunctionTest.text(1.134, "0,0").should == "1.1"
    FunctionTest.text(123456789.123456, "#.##").should == "123,456,789"
    FunctionTest.text(123456789.123456, "#.##0,0").should == "123,456,789.1"
  end

  it "should cope with weird combinations" do
    FunctionTest.text(3.1, "#,000").should == "003"
    FunctionTest.text(1000.3, "#,000").should == "1,000"
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
