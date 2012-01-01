require_relative '../spec_helper'

describe ExtractWorksheetDimensions do
  
  it "should output the dimensions of the worksheet" do
    input = excel_fragment 'ValueTypes.xml'
    output = StringIO.new
    ExtractWorksheetDimensions.extract(input,output)
    output.string.should == "A1:A6\n"
  end
end
