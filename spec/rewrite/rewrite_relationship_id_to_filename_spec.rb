require_relative '../spec_helper'

describe RewriteRelationshipIdToFilename do
  
  it "should take a relationship file and use it to map the last column of an input file from rIds to filenames" do
    input = StringIO.new("rId1\nrId2\nrId3\n")
    relationships = StringIO.new("rId1\tworksheets/sheet1.xml\nrId2\tworksheets/sheet2.xml\nrId3\tworksheets/sheet3.xml\nrId4\ttheme/theme1.xml\n")
    output = StringIO.new
expected = <<END
worksheets/sheet1.xml
worksheets/sheet2.xml
worksheets/sheet3.xml
END
    RewriteRelationshipIdToFilename.rewrite(input,relationships,output)
    output.string.should == expected
  end
end
