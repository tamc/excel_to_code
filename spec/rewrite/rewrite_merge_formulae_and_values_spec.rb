require_relative '../spec_helper'

describe RewriteMergeFormulaeAndValues do
  
it "should take files of the different types of formulae and a file of values and combine them into a single file" do

values = <<END
A1	[:boolean_true]
A2	[:string, "Hello"]
A3	[:number, "1"]
A4	[:number, "3.1415000000000002"]
A5	[:error, "#NAME?"]
A6	[:string, "Hello"]
END

simple_formulae = <<END
A1	[:function,"IF"]
A4	[:function,"IF"]
END

array_formulae = <<END
A2	[:function,"INDEX"]
END

shared_formulae = <<END
A3	[:function,"MATCH"]
A4	[:function,"MATCH"]
END

expected_output = <<END
A1	[:function,"IF"]
A2	[:function,"INDEX"]
A3	[:function,"MATCH"]
A4	[:function,"IF"]
A5	[:error, "#NAME?"]
A6	[:string, "Hello"]
B1	[:blank]
END

output = StringIO.new
r = RewriteMergeFormulaeAndValues.new
r.references_to_add_if_they_are_not_already_present = ["A1","B1"]
r.rewrite(StringIO.new(values),StringIO.new(shared_formulae),StringIO.new(array_formulae),StringIO.new(simple_formulae),output)
output.string.should == expected_output
end
end
