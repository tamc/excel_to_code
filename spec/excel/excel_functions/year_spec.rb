require_relative '../../spec_helper.rb'

describe "ExcelFunctions: YEAR" do
  
  it "should return something when given appropriate arguments" do
    File.open("../../test_data/dateinput.csv") do |file|
      file.each do |line|
        arr = line.split(",")
        FunctionTest.year(arr[0].to_i).should == arr[3].to_i
      end
    end
  end

  it "should return an error when given inappropriate arguments" do
    FunctionTest.year("Asdasddf").should == :value
  end
    
  it "should treat nil as an error" do
    FunctionTest.year(nil).should == :error
  end
  
  it "should return an error if an argument is an error" do
    FunctionTest.year(:error).should == :error
  end
  
  it "should be in the list of functions that can be mapped to ruby" do
    MapFormulaeToRuby::FUNCTIONS[:'YEAR'].should == 'year'
  end
  
end
