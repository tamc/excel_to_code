require_relative '../../spec_helper.rb'


describe "ExcelFunctions: MONTH" do
  
  it "should return something when given appropriate arguments" do
    File.open("../../test_data/dateinput.csv") do |file|
        file.each do |line|
          arr = line.split(",")
          FunctionTest.month(arr[0].to_i).should == arr[2].to_i
        end
      end
  end

  it "should return an error when given inappropriate arguments" do
    FunctionTest.month("Asdasddf").should == :value
  end
    
  it "should treat nil as zero" do
    FunctionTest.month(nil).should == 0
  end
  
  it "should return an error if an argument is an error" do
    FunctionTest.month(:error).should == :error
  end
  
  it "should be in the list of functions that can be mapped to ruby" do
    MapFormulaeToRuby::FUNCTIONS[:'MONTH'].should == 'month'
  end
  
end
