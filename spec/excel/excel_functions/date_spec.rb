require_relative '../../spec_helper.rb'


describe "ExcelFunctions: DATE" do
  
  it "should return a valid date when given a valid year, month, day" do
    File.open("../../test_data/dateinput.csv") do |file|
      file.each do |line|
        arr = line.split(",")
        FunctionTest.date(arr[3].to_i, arr[2].to_i, arr[1].to_i).should == arr[0].to_i
      end
    end
  end

  it "should return an error when given inappropriate arguments" do
    FunctionTest.date("Asdasddf").should == :num
  end
    
  it "should treat nil as an error" do
    FunctionTest.date(nil).should == :num
  end
  
  it "should return an error if an argument is an error" do
    FunctionTest.date(:error).should == :error
  end
  
  it "should be in the list of functions that can be mapped to ruby" do
    MapFormulaeToRuby::FUNCTIONS[:'DATE'].should == 'date'
  end
  
end