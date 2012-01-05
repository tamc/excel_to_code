require_relative '../spec_helper'

describe ExtractWorksheetTableRelationships do
  
  it "should output the worksheet names from the workbook, together with ids" do
    input = excel_fragment 'TableRelationships.xml'
    output = StringIO.new
expected = <<END
rId1
rId2
rId3
END
    ExtractWorksheetTableRelationships.extract(input,output)
    output.string.should == expected
  end
end
