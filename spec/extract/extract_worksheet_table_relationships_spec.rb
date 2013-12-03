require_relative '../spec_helper'

describe ExtractWorksheetTableRelationships do
  
  it "should output the worksheet names from the workbook, together with ids" do
    input = excel_fragment 'TableRelationships.xml'
    output = ExtractWorksheetTableRelationships.extract(input)
    output.should == ["rId1", "rId2", "rId3"]
  end
end
