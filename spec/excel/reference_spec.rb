require_relative '../spec_helper'
require_relative '../../src/excel/reference'

describe Reference do
  it "should take excel references as its initializer and return excel references when called with to_s" do
    Reference.new("A1").to_s.should == "A1"
    Reference.new("AAA$305").to_s.should == "AAA$305"
  end
  
  it "should be able to offset a reference" do
    Reference.new("A1").offset(0,0).to_s.should == "A1"
    Reference.new("A1").offset(1,1).to_s.should == "B2"
    Reference.new("Z1").offset(1,1).to_s.should == "AA2"
  end
end