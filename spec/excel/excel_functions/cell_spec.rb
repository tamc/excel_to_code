require_relative '../../spec_helper.rb'

describe "ExcelFunctions: CELL" do
  
  it "should return the filename of the original excel file when passed 'filename' as its first argument" do
    FunctionTest.cell('filename', nil).should == "filename not specified"
  end

  it "should return an error when given inappropriate arguments" do
    FunctionTest.cell("Asdasddf").should == :value
  end
    
  it "should return an error if an argument is an error" do
    FunctionTest.cell(:error).should == :error
  end
  
  it "should be in the list of functions that can be mapped to ruby" do
    MapFormulaeToRuby::FUNCTIONS[:'CELL'].should == 'cell'
  end
  
end
