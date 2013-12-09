require_relative '../../spec_helper.rb'

describe "ExcelFunctions: FIND(text,within_text,[start_num])" do
  
  it "should find the first occurrence of one string in another, returning :value if the string doesn't match'" do
    FunctionTest.find("one","onetwothree").should == 1
    FunctionTest.find("one","twoonethree").should == 4
    FunctionTest.find("one","twoonthree").should == :value
  end
  
  it "should find the first occurrence of one string in another after a given index, returning :value if the string doesn't match" do
    FunctionTest.find("one","onetwothree",1).should == 1
    FunctionTest.find("one","twoonethree",5).should == :value
    FunctionTest.find("one","oneone",2).should == 4
  end
  
  it "should be possible for the start_num to be a string, if that string converts to a number" do
    FunctionTest.find("one","oneone","2").should == 4
  end
  
  it "should return a :value error when given anything but a number as the third argument" do
    FunctionTest.find("one","oneone","a").should == :value
  end
  
  it "should return a :value error when given a third argument that is less than 1 or greater than the length of the string" do
    FunctionTest.find("one","oneone",0).should == :value
    FunctionTest.find("one","oneone",-1).should == :value
    FunctionTest.find("one","oneone",7).should == :value
  end
  
  it "nil in the first argument matches any character" do
    FunctionTest.find(nil,"abcdefg").should == 1
    FunctionTest.find(nil,"abcdefg",4).should == 4
  end

  it "should treat nil in the second argument as an empty string" do
    FunctionTest.find(nil,nil).should == 1
    FunctionTest.find("a",nil).should == :value
  end
    
  it "should return an error if any argument is an error" do
    FunctionTest.find("one","onetwothree",:error3).should == :error3
    FunctionTest.find("one",:error2,1).should == :error2
    FunctionTest.find(:error1,"onetwothree",1).should == :error1
    FunctionTest.find(:error1,:error2,:error3).should == :error1
  end
  
  it "should be in the list of functions that can be mapped to ruby" do
    MapFormulaeToRuby::FUNCTIONS[:'FIND'].should == 'find'
  end
  
end
