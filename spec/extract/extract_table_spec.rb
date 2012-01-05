require_relative '../spec_helper'

describe ExtractTable do
  
it "should create a flat file with one string per table, in the format name, sheet, range, number of total rows, column names" do
input = excel_fragment 'Table.xml'
expected_output = <<END
FirstTable	Sheet1	B2:C5	1	ColA	ColB
END

output = StringIO.new
ExtractTable.new('Sheet1').extract(input,output)
output.string.should == expected_output
end
end
