require_relative '../spec_helper'

describe "Object#try" do
  
  it "should return the value of the method if the object supports the method" do
    [1,2].try(:last).should == 2
  end
  
  it "should return nil otherwise" do
    nil.try(:last).should == nil
    1.try(:last).should == nil
  end
    
end
