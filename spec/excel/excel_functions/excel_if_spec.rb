require_relative '../../spec_helper.rb'
include ExcelFunctions

describe "ExcelFunctions: IF() excel_if()" do
  
  it "should return its second argument if the first argument is true" do
    excel_if(true,1,0).should == 1
  end

  it "should return its third argument if the first argument is false" do
    excel_if(false,1,0).should == 0
  end

  it "the third argument is optional, it will return false if it isn't specified" do
    excel_if(false,1).should == false
  end
  
end
