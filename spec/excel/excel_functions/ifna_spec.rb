require_relative '../../spec_helper.rb'

describe "ExcelFunctions: IFNA" do
  
  it "should return its first argument if it is not an NA" do
    FunctionTest.ifna(1, 2).should == 1
  end

  it "should return its second argument if the first is an NA error" do
    FunctionTest.ifna(:na, 2).should == 2
  end

  it "should return an error if its first argument is any other error" do
    FunctionTest.ifna(:error, 2).should == :error
  end
    
  it "should treat nil as zero" do
    FunctionTest.ifna(nil, 1).should == 0
    FunctionTest.ifna(:na, nil).should == 0
  end
  
  it "should be in the list of functions that can be mapped to ruby" do
    MapFormulaeToRuby::FUNCTIONS[:'IFNA'].should == 'ifna'
  end
  
end
