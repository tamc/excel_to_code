require_relative '../../spec_helper.rb'


describe "ExcelFunctions: DAY" do
  
  it "should return a valid DAY when given a sequence number" do
    File.open("../../test_data/dateinput.csv") do |file|
      file.each do |line|
        arr = line.split(",")
        FunctionTest.day(arr[0].to_i).should == arr[1].to_i
      end
    end
  end

  it "should return an error when given inappropriate arguments" do
    FunctionTest.day("Asdasddf").should == :value
  end
    
  it "should treat nil as an error" do
    FunctionTest.day(nil).should == 0
  end
  
  it "should return an error if an argument is an error" do
    FunctionTest.day(:error).should == :error
  end
  
  it "should be in the list of functions that can be mapped to ruby" do
    MapFormulaeToRuby::FUNCTIONS[:'DAY'].should == 'day'
  end
  
end
