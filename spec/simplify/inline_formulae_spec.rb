require_relative '../spec_helper'

describe InlineFormulae do
  
it "should recursively work through formulae, inlining references" do

input = <<END
A1\t[:cell, "$A$2"]
A2\t[:cell, "A3"]
A3\t[:number, 1]
A4\t[:sheet_reference,"sheet2",[:cell,"A1"]]
END

references = references = {
  'sheet1' => {
    'A1' => [:cell, "$A$2"],
    'A2' => [:cell, "A3"],
    'A3' => [:number, 1]
  },
  'sheet2' => {
    'A1' => [:cell, "A2"],
    'A2' => [:sheet_reference,'sheet3',[:cell,'A1']]
  },
  'sheet3' => {
    'A1' => [:number, 5]
  }
}

expected_output = <<END
A1\t[:number, 1]
A2\t[:number, 1]
A3\t[:number, 1]
A4\t[:number, 5]
END
    
input = StringIO.new(input)
output = StringIO.new
r = InlineFormulae.new
r.references = references
r.default_sheet_name = 'sheet1'
r.replace(input,output)
output.string.should == expected_output
end
end
