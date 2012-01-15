require_relative '../spec_helper'

describe ReplaceBlanks do
  
  it "should replace references that point to empty cells with [:blank]" do

input = <<END
A1\t[:cell, "B1"]
A2\t[:cell, "C1"]
A3\t[:sheet_reference, "sheet2", [:cell, "B1"]]
A4\t[:sheet_reference, "sheet2", [:cell, "C1"]]
END

references = {
  'sheet1' => {
    'A1' => true,
    'A2' => true,
    'B1' => true
  },
  'sheet2' => {
    'C1' => true
  }
}

expected_output = <<END
A1\t[:cell, "B1"]
A2\t[:blank]
A3\t[:blank]
A4\t[:sheet_reference, "sheet2", [:cell, "C1"]]
END
    
input = StringIO.new(input)
output = StringIO.new
ReplaceBlanks.replace(input,references,'sheet1',output)
output.string.should == expected_output
end
end
