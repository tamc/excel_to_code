require_relative '../../spec_helper.rb'

describe "ExcelFunctions: FORECAST" do
  
  it "should return a value based on the linear fit of the passed values" do
    known_x = [[1,2,3,4,5]]
    known_y = [[2,3,4,5,6]]
    FunctionTest.forecast(0,known_y, known_x).should == 1
    FunctionTest.forecast(1,known_y, known_x).should == 2
    FunctionTest.forecast(6,known_y, known_x).should == 7
  end

  it "should return an error when given inappropriate arguments" do
    FunctionTest.forecast(0,[[]], [[]]).should == :na
    FunctionTest.forecast(0,[[1,2,3]], [[1,2]]).should == :na
    FunctionTest.forecast(0,[[1,2]], [[1,2,3]]).should == :na
    FunctionTest.forecast("Asdasddf",[[1,2]],[[1,2]]).should == :value
  end
    
  it "should treat nil as zero when looking for a value, and ignore them when forecasting" do
    known_x = [[1,2,3,4,5]]
    known_y = [[2,3,4,5,6]]
    FunctionTest.forecast(nil, known_y, known_x).should == 1
    known_x = [[nil,2,3,4,nil]]
    FunctionTest.forecast(6, known_y, known_x).should == 7
  end
  
  it "should return an error if an argument is an error" do
    known_x = [[1,2,3,4,5]]
    known_y = [[2,3,4,5,6]]
    FunctionTest.forecast(:error,known_y, known_x).should == :error
    known_x[1] = :error
    FunctionTest.forecast(6,known_y, known_x).should == :error
  end
  
  it "should be in the list of functions that can be mapped to ruby" do
    MapFormulaeToRuby::FUNCTIONS[:'FORECAST'].should == 'forecast'
  end
  
end
