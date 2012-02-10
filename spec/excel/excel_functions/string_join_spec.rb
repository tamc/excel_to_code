require_relative '../../spec_helper.rb'

describe "ExcelFunctions: string_join(string,string)" do
  
  it "should return a string by combining its arguments" do
    FunctionTest.string_join("Hello ","world").should == "Hello world"
  end
  
  it "should cope with an arbitrary number of arguments" do
    FunctionTest.string_join("Hello"," ","world","!").should == "Hello world!"
  end
    
  it "should convert values to strings as it goes" do
    FunctionTest.string_join("Top ",10).should == "Top 10"
  end
  
  it "should convert integer values into strings without decimal points" do
    FunctionTest.string_join("Top ",10.0).should == "Top 10"
    FunctionTest.string_join("Top ",10.5).should == "Top 10.5"
  end
  
  
  it "should return an error if an argument is an error" do
    FunctionTest.string_join(:error,1).should == :error
    FunctionTest.string_join(1,:error).should == :error
    FunctionTest.string_join(:error1,:error2).should == :error1
  end
  
end
