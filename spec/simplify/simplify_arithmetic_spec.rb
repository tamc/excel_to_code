require_relative '../spec_helper'

describe SimplifyArithmetic do
  
it "should turn [:arithmetic] ast with more than one operation into a series of nested arithmetic statements, following the excel precedence rules" do

input = <<END
A1\t[:arithmetic, [:number, "1"], [:operator, "+"], [:number, "2"], [:operator, "+"], [:number, "3"]]
A2\t[:arithmetic, [:number, "1"], [:operator, "+"], [:number, "2"], [:operator, "-"], [:number, "3"]]
A3\t[:arithmetic, [:number, "1"], [:operator, "/"], [:number, "2"], [:operator, "*"], [:number, "3"]]
A4\t[:arithmetic, [:number, "1"], [:operator, "*"], [:number, "2"], [:operator, "^"], [:number, "3"]]
A5\t[:arithmetic, [:number, "1"], [:operator, "/"], [:number, "2"], [:operator, "^"], [:number, "3"]]
A6\t[:arithmetic, [:number, "1"], [:operator, "+"], [:number, "2"], [:operator, "*"], [:number, "3"]]
A7\t[:arithmetic, [:number, "1"], [:operator, "+"], [:number, "2"], [:operator, "/"], [:number, "3"]]
A8\t[:arithmetic, [:number, "1"], [:operator, "-"], [:number, "2"], [:operator, "*"], [:number, "3"]]
A9\t[:arithmetic, [:number, "1"], [:operator, "-"], [:number, "2"], [:operator, "/"], [:number, "3"]]
A10\t[:arithmetic, [:number, "1"], [:operator, "-"], [:number, "2"], [:operator, "/"], [:number, "3"],  [:operator, "+"], [:number, "4"]]
END

expected_output = <<END
A1\t[:arithmetic, [:arithmetic, [:number, "1"], [:operator, "+"], [:number, "2"]], [:operator, "+"], [:number, "3"]]
A2\t[:arithmetic, [:arithmetic, [:number, "1"], [:operator, "+"], [:number, "2"]], [:operator, "-"], [:number, "3"]]
A3\t[:arithmetic, [:arithmetic, [:number, "1"], [:operator, "/"], [:number, "2"]], [:operator, "*"], [:number, "3"]]
A4\t[:arithmetic, [:number, "1"], [:operator, "*"], [:arithmetic, [:number, "2"], [:operator, "^"], [:number, "3"]]]
A5\t[:arithmetic, [:number, "1"], [:operator, "/"], [:arithmetic, [:number, "2"], [:operator, "^"], [:number, "3"]]]
A6\t[:arithmetic, [:number, "1"], [:operator, "+"], [:arithmetic, [:number, "2"], [:operator, "*"], [:number, "3"]]]
A7\t[:arithmetic, [:number, "1"], [:operator, "+"], [:arithmetic, [:number, "2"], [:operator, "/"], [:number, "3"]]]
A8\t[:arithmetic, [:number, "1"], [:operator, "-"], [:arithmetic, [:number, "2"], [:operator, "*"], [:number, "3"]]]
A9\t[:arithmetic, [:number, "1"], [:operator, "-"], [:arithmetic, [:number, "2"], [:operator, "/"], [:number, "3"]]]
A10\t[:arithmetic, [:arithmetic, [:number, "1"], [:operator, "-"], [:arithmetic, [:number, "2"], [:operator, "/"], [:number, "3"]]], [:operator, "+"], [:number, "4"]]
END
    
input = StringIO.new(input)
output = StringIO.new
r = SimplifyArithmetic.new
r.replace(input,output)
output.string.should == expected_output

end # / it


end # / describe
