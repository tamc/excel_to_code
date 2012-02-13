require_relative '../spec_helper'

describe ReplaceCommonElementsInFormulae do
  
it "should work through formulae, replacing elements that are common to other formulae with a cell reference" do

input = <<END
A1\t[:function, "INDEX", [:array, [:row, [:cell, "A1"], [:cell, "A2"], [:cell, "A3"]]], [:number, 2]]
END

common = <<END
common0\t[:array, [:row, [:cell, "A1"], [:cell, "A2"], [:cell, "A3"]]]
END

expected_output = <<END
A1\t[:function, "INDEX", [:cell, "common0"], [:number, 2]]
common0\t[:array, [:row, [:cell, "A1"], [:cell, "A2"], [:cell, "A3"]]]
END
    
input = StringIO.new(input)
common = StringIO.new(common)
output = StringIO.new
r = ReplaceCommonElementsInFormulae.new
r.replace(input,common,output)
output.string.should == expected_output
end # /it

end # /Describe
