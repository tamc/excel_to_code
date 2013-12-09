require_relative '../spec_helper'

describe ReplaceNamedReferences do
  
  it "should replace named references with the references that they point to" do

input = <<END
A1\t[:named_reference, "Global"]
A2\t[:named_reference, "Local"]
A3\t[:sheet_reference,"otherSheet",[:named_reference, "Local"]]
A4\t[:sheet_reference,"otherSheet",[:named_reference, "Local"]]
A5\t[:named_reference, "missing"]
END

named_references = {
  "global" =>               [:sheet_reference,'thisSheet',[:area, "A1:A10"]], 
  "local" =>                [:sheet_reference,'notReallyLocal',[:area, "A1:A10"]],
  ["thissheet","local"] =>  [:sheet_reference,'thisSheet',[:area, "A1:A10"]],
  ["othersheet","local"] => [:sheet_reference,'otherSheet',[:area, "A1:A10"]]
}

expected_output = <<END
A1\t[:sheet_reference, "thisSheet", [:area, "A1:A10"]]
A2\t[:sheet_reference, "thisSheet", [:area, "A1:A10"]]
A3\t[:sheet_reference, "otherSheet", [:area, "A1:A10"]]
A4\t[:sheet_reference, "otherSheet", [:area, "A1:A10"]]
A5\t[:error, :"#NAME?"]
END
    
input = StringIO.new(input)
output = StringIO.new
r = ReplaceNamedReferences.new
r.sheet_name = "thisSheet"
r.named_references = named_references
r.replace(input,output)
output.string.should == expected_output
end

it "should work even if no local references" do

input = <<END
A1\t[:named_reference, "gLOBal"]
END

named_references = {"global" => [:sheet_reference,'thisSheet',[:area, "A1:A10"]]}

expected_output = <<END
A1\t[:sheet_reference, "thisSheet", [:area, "A1:A10"]]
END
    
input = StringIO.new(input)
output = StringIO.new
r = ReplaceNamedReferences.new
r.sheet_name = "thisSheet"
r.named_references = named_references
r.replace(input,output)
output.string.should == expected_output
end

it "should work for array references" do

input = <<END
A1\tA1:B6\t[:named_reference, "Global"]
END

named_references = {"global" => [:sheet_reference,'thisSheet',[:area, "A1:A10"]]}

expected_output = <<END
A1\tA1:B6\t[:sheet_reference, "thisSheet", [:area, "A1:A10"]]
END
    
input = StringIO.new(input)
output = StringIO.new
r = ReplaceNamedReferences.new
r.sheet_name = "thisSheet"
r.named_references = named_references
r.replace(input,output)
output.string.should == expected_output
end



end
