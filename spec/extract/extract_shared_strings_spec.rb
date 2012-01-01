require_relative '../spec_helper'

describe ExtractSharedStrings do
  
  it "should create a flat file with one string per row" do
    input = excel_fragment 'SharedStrings.xml'
    output = StringIO.new
    ExtractSharedStrings.extract(input,output)
    output.string.should == "This a second shared string\nThis is, hopefully, the first shared string\n"
  end
end
