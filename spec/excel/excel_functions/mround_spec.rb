require_relative '../../spec_helper.rb'

describe "ExcelFunctions: MROUND" do
  
  it "should return something when given appropriate arguments" do
    [
      [nil, nil, 0],
      [1, nil, 0],
      [nil, 1, 0],
      ["a",1, :value],
      [1, "a", :value],
      [:error, 1, :error],
      [1, :error, :error],
      [1, 1, 1],
      [105, 10, 110],
      [1.05, 0.1, 1.1],
      [-1.05, -0.1, -1.1],
      [1.5, -0.1, :num],
    ].each do |t|
      a, b, expected = t
      FunctionTest.mround(a, b).should == expected
    end
  end
  
  it "should be in the list of functions that can be mapped to ruby" do
    MapFormulaeToRuby::FUNCTIONS[:'MROUND'].should == 'mround'
  end
  
end
