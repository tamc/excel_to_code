require_relative '../spec_helper'

describe ExtractSharedStrings do
  
  it "should create a flat file with one string per row" do
    input = excel_fragment 'SharedStrings.xml'
    output = ExtractSharedStrings.extract(input)
    output.should == [ [:string, "This a second shared string"], [:string, "This is, hopefully, the first shared string"] ]
  end
end
