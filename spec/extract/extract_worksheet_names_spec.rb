require_relative '../spec_helper'

describe ExtractWorksheetNames do
  
  it "should output the worksheet names from the workbook, together with ids" do
    input = excel_fragment 'Workbook.xml'
    output = ExtractWorksheetNames.extract(input)
    output.should == {"Outputs" => "rId1" , "Calcs" => "rId2", "Inputs" => "rId3"}
  end
end
