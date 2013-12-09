require_relative '../spec_helper'

describe Reference do
  it "should have a Reference#for method that takes a reference string" do
    Reference.for("A1").should == "A1"
    Reference.for("AAA$305").should == "AAA$305"
  end
  
  it "should be a subclass of String" do
    Reference.for("A1").should be_kind_of(String)
  end

  it "should work for symbols" do
    Reference.for(:A1).should be_kind_of(String)
  end
  
  it "should return the same instance each time a particular reference is required" do
    Reference.for("A1").object_id.should ==  Reference.for("A1").object_id
  end
  
  it "should be able to offset a reference" do
    Reference.for("A1").offset(0,0).should == :A1
    Reference.for("A1").offset(1,1).should == :B2
    Reference.for("Z1").offset(1,1).should == :AA2
  end
  
  it "should respect fixed reference when offseting" do
    Reference.for("$A$1").offset(1,1).should == :"$A$1"
    Reference.for("$A1").offset(1,1).should == :"$A2"
    Reference.for("A$1").offset(1,1).should == :"B$1"
  end
  
  it "should be able to unfix a reference" do 
    Reference.for("A1").unfix.should == "A1"
    Reference.for("$A$1").unfix.should == "A1"
    Reference.for("$A1").unfix.should == "A1"
    Reference.for("A$1").unfix.should == "A1"
  end
  
end
