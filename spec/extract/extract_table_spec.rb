require_relative '../spec_helper'

describe ExtractTable do
  
it "should create a hash: table_name => [sheet, range, number of total rows, column names]" do
input = excel_fragment 'Table.xml'
expected_output = { "FirstTable" =>	["Sheet1", "B2:C5", "1", "ColA", "ColB"] }

output = ExtractTable.extract("Sheet1", input)
output.should == expected_output
end

it "should work even when the table does not have a total row" do
input = excel_fragment 'Table2.xml'
expected_output = { "FirstTable" =>	["Sheet1", "B2:C5", "0", "ColA", "ColB"] }

output = ExtractTable.extract("Sheet1", input)
output.should == expected_output
end


end
