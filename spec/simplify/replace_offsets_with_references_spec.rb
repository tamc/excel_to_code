require_relative '../spec_helper'

describe ReplaceOffsetsWithReferences do
  
it "should replace OFFSET() functions with the reference that they refer to, if they have been passed values" do

input = <<END
A1\t[:function, :OFFSET, [:cell, :"$A$5"], [:number, 1], [:number, 1], [:number, 3], [:number, 3]] 
A2\t[:function, :OFFSET, [:cell, :"$A$5"], [:number, 1], [:number, 1]] 
A3\t[:function, :OFFSET, [:cell, :"$A$5"], [:cell, :Z10], [:number, 1], [:number, 3], [:number, 3]] 
A4\t[:string_join, [:string, "Chosen language is"], [:string, " "], [:function, :OFFSET, [:cell, :"B18"], [:number, "0"], [:number, "0"]]]
A5\t[:function, :OFFSET, [:sheet_reference, "User inputs", [:cell, :"$D$44"]], [:number, "0"], [:number, "-1"]]
END

expected_output = <<END
A1\t[:area, :B6, :D8]
A2\t[:cell, :B6]
A3\t[:function, :OFFSET, [:cell, :"$A$5"], [:cell, :Z10], [:number, 1], [:number, 3], [:number, 3]]
A4\t[:string_join, [:string, "Chosen language is"], [:string, " "], [:cell, :B18]]
A5\t[:sheet_reference, "User inputs", [:cell, :C44]]
END
    
input = StringIO.new(input)
output = StringIO.new
r = ReplaceOffsetsWithReferences.new
r.replace(input,output)
output.string.should == expected_output

end # / it


end # / describe
