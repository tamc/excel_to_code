require_relative '../../spec_helper.rb'

describe "ExcelFunctions: LOOKUP" do
  
  it "should return the value in the result vector that matches the index of lookup_value in lookup_vector" do
    lookup_vector = [
      [ 4.14 ],
      [ 4.19 ],
      [ 5.17 ],
      [ 5.77 ],
      [ 6.39 ],
    ]
    result_vector = [
      [ "red" ],
      [ "orange" ],
      [ "yellow" ],
      [ "green" ],
      [ "blue" ],
    ]
    FunctionTest.lookup(4.19, lookup_vector, result_vector).should == "orange"
    FunctionTest.lookup(5.75, lookup_vector, result_vector).should == "yellow"
    FunctionTest.lookup(7.66, lookup_vector, result_vector).should == "blue"
    FunctionTest.lookup(7.66, lookup_vector).should == 6.39
    FunctionTest.lookup(0, lookup_vector, result_vector).should == :na
  end

  it "should work in array form as well" do
    lookup_array = [
      [ 4.14, "red" ],
      [ 4.19, "orange" ],
      [ 5.17, "yellow" ],
      [ 5.77, "green" ],
      [ 6.39, "blue" ],
    ]
    FunctionTest.lookup(4.19, lookup_array).should == "orange"
    FunctionTest.lookup(5.75, lookup_array).should == "yellow"
    FunctionTest.lookup(7.66, lookup_array).should == "blue"
    FunctionTest.lookup(0, lookup_array).should == :na
  end

  it "should treat nil as zero" do
    FunctionTest.lookup(nil, [[0, 1, 2]], [["a", "b", "c"]]).should == "a"
  end
  
  it "should return an error if an argument is an error" do
    FunctionTest.lookup(:error, [[1,2,3]]).should == :error
    FunctionTest.lookup(1, :error).should == :error
    FunctionTest.lookup(1, [[1,2,3]], :error).should == :error
    # The below may be what Excel wants?
    # FunctionTest.lookup(2.4, [[1,:na,3]]).should == 1 
  end
  
  it "should be in the list of functions that can be mapped to ruby" do
    MapFormulaeToRuby::FUNCTIONS[:'LOOKUP'].should == 'lookup'
  end
  
end
