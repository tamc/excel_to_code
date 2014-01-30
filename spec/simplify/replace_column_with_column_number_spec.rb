require_relative '../spec_helper'

describe ReplaceColumnAndRowFunctions do
  
it "should replace COLUMN() and ROW() functions with the number of the column or row that they refer to" do

input = <<END
A1\t[:function, :COLUMN, [:cell, :"$A$5"]]
A2\t[:function, :COLUMN, [:sheet_reference, "Sheet1", [:cell, :G70]]]
A3\t[:function, :ROW, [:cell, :"$A$5"]]
A4\t[:function, :ROW, [:sheet_reference, "Sheet1", [:cell, :G70]]]
A5\t[:function, :ROW]
A6\t[:function, :COLUMN]
END

expected_output = <<END
A1\t[:number, 1.0]
A2\t[:number, 7.0]
A3\t[:number, 5.0]
A4\t[:number, 70.0]
A5\t[:number, 5.0]
A6\t[:number, 1.0]
END
    
input = StringIO.new(input)
output = StringIO.new
r = ReplaceColumnAndRowFunctions.new
r.replace(input,output)
output.string.should == expected_output

end # / it


end # / describe
