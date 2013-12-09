require_relative '../../spec_helper.rb'

describe "ExcelFunctions: IFERROR(value,value_if_error)" do
  
  it "should return its second value if there is an error in the first" do
    FunctionTest.iferror("ok","Not found").should == "ok"
    FunctionTest.iferror(:value,"Not found").should == "Not found"
  end
  
  it "should be able to trap divide by zero errors" do
    FunctionTest.iferror(0.0/0.0,"Zero division").should == "Zero division"
    FunctionTest.iferror(FunctionTest.divide(0,0),"Zero division").should == "Zero division"
  end
    
  it "should treat nil as zero" do
    FunctionTest.iferror(:error,nil).should == 0
  end
    
  it "should be in the list of functions that can be mapped to ruby" do
    MapFormulaeToRuby::FUNCTIONS[:'IFERROR'].should == 'iferror'
  end
  
end
