require_relative '../spec_helper'

describe ReplaceCellAddressesWithReferences do
  
it "should replace CELL('address',A1) functions with the reference that they refer to" do

input = <<END
A1\t[:function, :CELL, [:string, "address"], [:cell, :A5]]
A2\t[:function, :CELL, [:string, "ADDRESS"], [:cell, :"$A$5"]]
A3\t[:function, :CELL, [:string, "ADDRESS"], [:sheet_reference, "Sheet1", [:cell, :"$A$5"]]]
A4\t[:function, :CELL, [:string, "address"], [:area, :A5, :B10]]
A5\t[:function, :CELL, [:string, "ADDRESS"], [:sheet_reference, "Sheet1", [:area, :"$A$5", :B10]]]
A6\t[:function, :CELL, [:string, "address"], [:number, 1]]
A7\t[:function, :CELL, [:string, "address"], [:string, "Hello"]]
A8\t[:function, :CELL, [:string, "address"], [:error, "#DIV/0!"]]
END

expected_output = <<END
A1\t[:cell, "A5"]
A2\t[:cell, "A5"]
A3\t[:sheet_reference, "Sheet1", [:cell, "A5"]]
A4\t[:cell, "A5"]
A5\t[:sheet_reference, "Sheet1", [:cell, "A5"]]
A6\t[:error, "#VALUE!"]
A7\t[:error, "#VALUE!"]
A8\t[:error, "#VALUE!"]
END
    
input = StringIO.new(input)
output = StringIO.new
r = ReplaceCellAddressesWithReferences.new
r.replace(input,output)
output.string.should == expected_output

end # / it


end # / describe
