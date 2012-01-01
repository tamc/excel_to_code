require_relative '../spec_helper'
require_relative '../../src/extract/extract_worksheet_dimensions'
require 'stringio'

describe ExtractWorksheetDimensions do
  
  it "should output the dimensions of the worksheet" do
    input = excel_fragment 'ValueTypes.xml'
    output = StringIO.new
    ExtractWorksheetDimensions.extract(input,output)
    output.string.should == "A1:A6\n"
  end
end
