require_relative '../spec_helper'

describe ReplaceColumnWithColumnNumber do
  
it "should replace COLUMN() functions with the number of the column that they refer to, if they have been passed values" do

input = <<END
A1\t[:function, :COLUMN, [:cell, :"$A$5"]]
A2\t[:function, :COLUMN, [:sheet_reference, "Sheet1", [:cell, :G70]]]
END

expected_output = <<END
A1\t[:number, 1.0]
A2\t[:number, 7.0]
END
    
input = StringIO.new(input)
output = StringIO.new
r = ReplaceColumnWithColumnNumber.new
r.replace(input,output)
output.string.should == expected_output

end # / it


end # / describe
