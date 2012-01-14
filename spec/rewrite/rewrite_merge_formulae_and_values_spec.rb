require_relative '../spec_helper'

describe RewriteMergeFormulaeAndValues do
  
it "should take a file of formulae and a file of values and output the two combined, but only including values where there isn't a formula'" do
values = <<END
A1	[:boolean_true]
A2	[:string, "Hello"]
A3	[:number, "1"]
A4	[:number, "3.1415000000000002"]
A5	[:error, "#NAME?"]
A6	[:string, "Hello"]
END

formulae = <<END
A1	[:function,"IF"]
A5	[:function,"IF"]
A6	[:function,"IF"]
END

expected_output = <<END
A1	[:function,"IF"]
A2	[:string, "Hello"]
A3	[:number, "1"]
A4	[:number, "3.1415000000000002"]
A5	[:function,"IF"]
A6	[:function,"IF"]
END

output = StringIO.new
RewriteMergeFormulaeAndValues.rewrite(StringIO.new(formulae),StringIO.new(values),output)
output.string.should == expected_output
end
end
