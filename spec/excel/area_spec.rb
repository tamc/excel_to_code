require_relative '../spec_helper'

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
  
  it "should be able to enumerate [row,col] offsets of references, relative to starting point" do
    Area.for("A1:B2").offsets.to_a == [[0,0],[1,0],[0,1],[1,1]]
  end
  
  it "should be able to return the height of the area" do
    Area.for("A1:B1").height.should == 0
    Area.for("A1:B2").height.should == 1
    Area.for("A1:A3").height.should == 2
  end
  
  it "should be able to return the width of the area" do
    Area.for("A1:A3").width.should == 0
    Area.for("A1:B3").width.should == 1
    Area.for("A3:C3").width.should == 2
  end
  
  it "should be able to return an equivalent array literal (e.g., {A1,B1;A2,B2})" do
    Area.for("A1:A1").to_array_literal.should == [:array,[:row,[:cell,'A1']]]
    Area.for("A1:A1").to_array_literal('worksheet').should == [:array,[:row,[:sheet_reference,'worksheet',[:cell,'A1']]]]
  end

end
