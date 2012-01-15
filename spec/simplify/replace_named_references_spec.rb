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

named_references = <<END
\tGlobal\t[:sheet_reference,'thisSheet',[:area, "A1:A10"]]
\tLocal\t[:sheet_reference,'notReallyLocal',[:area, "A1:A10"]]
thisSheet\tLocal\t[:sheet_reference,'thisSheet',[:area, "A1:A10"]]
otherSheet\tLocal\t[:sheet_reference,'otherSheet',[:area, "A1:A10"]]
END

expected_output = <<END
A1\t[:sheet_reference, "thisSheet", [:area, "A1:A10"]]
A2\t[:sheet_reference, "thisSheet", [:area, "A1:A10"]]
A3\t[:sheet_reference, "otherSheet", [:area, "A1:A10"]]
A4\t[:sheet_reference, "otherSheet", [:area, "A1:A10"]]
A5\t[:error, "#NAME?"]
END
    
    input = StringIO.new(input)
    named_references = StringIO.new(named_references)
    output = StringIO.new
    ReplaceNamedReferences.replace(input,"thisSheet",named_references,output)
    output.string.should == expected_output
  end

  it "should work even if no local references" do

input = <<END
A1\t[:named_reference, "Global"]
END

named_references = <<END
\tGlobal\t[:sheet_reference,'thisSheet',[:area, "A1:A10"]]
END

expected_output = <<END
A1\t[:sheet_reference, "thisSheet", [:area, "A1:A10"]]
END
    
    input = StringIO.new(input)
    named_references = StringIO.new(named_references)
    output = StringIO.new
    ReplaceNamedReferences.replace(input,"thisSheet",named_references,output)
    output.string.should == expected_output
  end


end
