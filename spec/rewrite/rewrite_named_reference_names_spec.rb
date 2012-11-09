require_relative '../spec_helper'

describe RewriteNamedReferenceNames do
  
it "should take the results of extract_named_references.rb and create references that are suitable for use in the C code" do

named_references = <<END
\tIn_result\tInputs!$A$3
\tXI.v.Scenario\tInputs!$A$4
Inputs\tLocal_named_reference\tInputs!$A$3
END

worksheet_names = <<END
Inputs\tinputs2
END
output = StringIO.new

RewriteNamedReferenceNames.rewrite(StringIO.new(named_references), StringIO.new(worksheet_names), output)

output.string.should == <<END
in_result\tInputs!$A$3
xi_v_scenario\tInputs!$A$4
inputs2_local_named_reference\tInputs!$A$3
END
end
end
