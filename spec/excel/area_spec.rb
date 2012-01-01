require_relative '../spec_helper'
require_relative '../../src/excel/area'

describe Area do
  it "should have a Area#for method that takes a reference string" do
    Area.for("A1:B10").should == "A1:B10"
    Area.for("AAA$305:$BBB1035").should == "AAA$305:$BBB1035"
  end
  
  it "should be a subclass of String" do
    Area.for("A1:B10").should be_kind_of(String)
  end
  
  it "should return the same instance each time a particular area is required" do
    Area.for("A1:B10").object_id.should ==  Area.for("A1:B10").object_id
  end
  
  it "should be able to offset an area" do
    Area.for("A1:B10").offset(0,0).should == "A1:B10"
    Area.for("A1:B10").offset(1,1).should == "B2:C11"
    Area.for("Z1:ZZ10").offset(1,1).should == "AA2:AAA11"
  end
  
  it "should respect fixed references when offseting" do
    Area.for("$A$1:B10").offset(1,1).should == "$A$1:C11"
    Area.for("$A1:B10").offset(1,1).should == "$A2:C11"
    Area.for("A$1:B10").offset(1,1).should == "B$1:C11"
  end
  
  it "should be able to enumerate offsets of references, relative to starting point" do
    Area.for("A1:B2").offsets.to_a == [[0,0],[1,0],[0,1],[1,1]]
  end

end
