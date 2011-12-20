require_relative '../spec_helper'
require_relative '../../src/extract/extract_relationships'
require 'stringio'

describe ExtractRelationships do
  
  it "should return a series of relationships" do
    input = excel_fragment 'Relationships.xml'
    output = StringIO.new
    ExtractRelationships.extract(input,output)
    output.string.should == "rId3\tworksheets/sheet3.xml\nrId4\ttheme/theme1.xml\n"
  end
end
