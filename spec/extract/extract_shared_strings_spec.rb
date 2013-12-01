require_relative '../spec_helper'

describe ExtractSharedStrings do
  
  it "should create a flat file with one string per row" do
    input = excel_fragment 'SharedStrings.xml'
    output = ExtractSharedStrings.extract(input)
    output.should == [ "This a second shared string", "This is, hopefully, the first shared string" ]
  end
end
