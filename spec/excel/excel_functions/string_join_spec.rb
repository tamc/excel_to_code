require_relative '../../spec_helper.rb'
include ExcelFunctions

describe "ExcelFunctions: string_join(string,string)" do
  
  it "should return a string by combining its arguments" do
    string_join("Hello ","world").should == "Hello world"
  end
  
  it "should cope with an arbitrary number of arguments" do
    string_join("Hello"," ","world","!").should == "Hello world!"
  end
    
  it "should convert values to strings as it goes" do
    string_join("Top ",10).should == "Top 10"
  end
  
  it "should return an error if an argument is an error" do
    string_join(:error,1).should == :error
    string_join(1,:error).should == :error
    string_join(:error1,:error2).should == :error1
  end
  
end
