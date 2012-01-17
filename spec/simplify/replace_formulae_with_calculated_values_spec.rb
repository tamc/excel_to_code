require_relative '../spec_helper'

describe ReplaceFormulaeWithCalculatedValues do
  
it "should work through formulae, calculating functions where all the arguments are given" do

input = <<END
A1\t[:arithmetic, [:number, "1"], [:operator, "+"], [:number, "1"]]
END

expected_output = <<END
A1\t[:number, "2"]
END
    
input = StringIO.new(input)
output = StringIO.new
r = ReplaceFormulaeWithCalculatedValues.new
r.replace(input,output)
output.string.should == expected_output
end # /it

end # /Describe
