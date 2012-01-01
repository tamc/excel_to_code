require_relative '../spec_helper'

describe RewriteWorksheetNames do
  
  it "should take the results of extract_worksheet_names.rb and extract_relationships.rb and return one line per worksheet: worksheet name then tab then worksheet filename" do
    extract_worksheet_names_output = StringIO.new("rId1\tOutputs\nrId2\tCalcs\nrId3\tInputs\n")
    extract_relationships_output = StringIO.new("rId1\tworksheets/sheet1.xml\nrId2\tworksheets/sheet2.xml\nrId3\tworksheets/sheet3.xml\nrId4\ttheme/theme1.xml\n")
    output = StringIO.new
    RewriteWorksheetNames.rewrite(extract_worksheet_names_output,extract_relationships_output,output)
    output.string.should == "Outputs\tworksheets/sheet1.xml\nCalcs\tworksheets/sheet2.xml\nInputs\tworksheets/sheet3.xml\n"
  end
end
