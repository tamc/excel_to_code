require_relative '../spec_helper'

describe ReplaceFormulaeWithCalculatedValues do
  
it "should work through formulae, calculating functions where all the arguments are given" do

input = <<END
A1\t[:arithmetic, [:number, 1], [:operator, :+], [:number, 1]]
A2\t[:function, :COUNT, [:array, [:row, [:cell, "B225"]], [:row, [:cell, "B226"]], [:row, [:cell, "B227"]], [:row, [:cell, "B228"]], [:row, [:cell, "B229"]], [:row, [:cell, "B230"]]]]
A3\t[:function, :SUMIFS, [:function, :INDEX, [:array, [:row, [:number, 1], [:number, 1], [:number, 0], [:number, 0]], [:row, [:number, 0], [:number, 0], [:number, 1], [:number, 0]], [:row, [:number, 0], [:number, 0], [:number, 0], [:number, 1]], [:row, [:number, 0], [:number, 0], [:number, 0], [:number, 0]], [:row, [:number, 0], [:number, 0], [:number, 1], [:number, 0]], [:row, [:number, 1], [:number, 1], [:number, 0], [:number, 1]], [:row, [:number, 0], [:number, 1], [:number, 0], [:number, 0]], [:row, [:number, 1], [:number, 0], [:number, 1], [:number, 0]], [:row, [:number, 0], [:number, 0], [:number, 0], [:number, 1]], [:row, [:number, 0], [:number, 0], [:number, 0], [:number, 0]], [:row, [:number, 1], [:number, 1], [:number, 1], [:number, 1]], [:row, [:number, 0], [:number, 0], [:number, 0], [:number, 0]], [:row, [:number, 0], [:number, 0], [:number, 0], [:number, 0]], [:row, [:number, 0], [:number, 0], [:number, 0], [:number, 0]], [:row, [:number, 1], [:number, 1], [:number, 1], [:number, 1]]], [:null], [:function, :MATCH, [:number, 1], [:array, [:row, [:number, 1], [:number, "2"], [:number, "3"], [:number, "4"]]], [:number, 0]]], [:array, [:row, [:string, "V.09"]], [:row, [:string, "V.09"]], [:row, [:string, "V.09"]], [:row, [:string, "V.10"]], [:row, [:string, "V.10"]], [:row, [:string, "V.10"]], [:row, [:string, "V.13"]], [:row, [:string, "V.13"]], [:row, [:string, "V.13"]], [:row, [:string, "V.14"]], [:row, [:string, "V.14"]], [:row, [:string, "V.14"]], [:row, [:string, "V.15"]], [:row, [:string, "V.15"]], [:row, [:string, "V.15"]]], [:string, "V.09"], [:array, [:row, [:string, "Solid"]], [:row, [:string, "Liquid"]], [:row, [:string, "Gas"]], [:row, [:string, "Solid"]], [:row, [:string, "Liquid"]], [:row, [:string, "Gas"]], [:row, [:string, "Solid"]], [:row, [:string, "Liquid"]], [:row, [:string, "Gas"]], [:row, [:string, "Solid"]], [:row, [:string, "Liquid"]], [:row, [:string, "Gas"]], [:row, [:string, "Solid"]], [:row, [:string, "Liquid"]], [:row, [:string, "Gas"]]], [:string, "Solid"]]
A4\t[:prefix, '-', [:boolean_false]]
A5\t[:prefix, '-', [:boolean_true]]
END

expected_output = <<END
A1\t[:number, 2.0]
A2\t[:number, 6.0]
A3\t[:number, 1.0]
A4\t[:number, 0.0]
A5\t[:number, -1.0]
END
    
input = StringIO.new(input)
output = StringIO.new
r = ReplaceFormulaeWithCalculatedValues.new
r.replace(input,output)
output.string.should == expected_output
end # /it

end # /Describe
