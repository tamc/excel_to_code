require_relative '../spec_helper'
require_relative '../../src/extract/extract_worksheet_names'
require 'stringio'

describe ExtractWorksheetNames do
  
  it "should output the worksheet names from the workbook, together with ids" do
    input = excel_fragment 'Workbook.xml'
    output = StringIO.new
    ExtractWorksheetNames.extract(input,output)
    output.string.should == "rId1\tOutputs\nrId2\tCalcs\nrId3\tInputs\n"
  end
end
