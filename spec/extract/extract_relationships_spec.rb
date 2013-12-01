require_relative '../spec_helper'

describe ExtractRelationships do
  
  it "should return a series of relationships" do
    input = excel_fragment 'Relationships.xml'
    output = ExtractRelationships.extract(input)
    output.should == { "rId3" => "worksheets/sheet3.xml", "rId4" => "theme/theme1.xml"}
  end
end
