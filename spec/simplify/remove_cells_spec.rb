require_relative '../spec_helper'

describe RemoveCells do
  
it "should remove any cells whose references are NOT in the given array" do
input = <<END
A1\t[:boolean_true]
A2\t[:shared_string, "0"]
A3\t[:number, "1"]
A4\t[:number, "3.1415000000000002"]
A5\t[:error, "#NAME?"]
A6\t[:string, "Hello    "]
END

cells_to_keep = {'A1' => true,'A2' => true, 'A6' => true}

expected_output = <<END
A1\t[:boolean_true]
A2\t[:shared_string, "0"]
A6\t[:string, "Hello    "]
END

output = StringIO.new
r = RemoveCells.new
r.cells_to_keep = cells_to_keep
r.rewrite(input,output)
output.string.should == expected_output

end # / it

end # /describe
