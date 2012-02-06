require_relative '../spec_helper'

describe RewriteArrayFormulaeToArrays do
  
  it "should take array formula like 'reference\\array range\\tformula ast\\n' and convert the formula into array form" do
    input = StringIO.new(%Q|B6\tB6:B8\t[:arithmetic, [:array, [:row, [:cell, "A1"]], [:row, [:cell, "A2"]]], [:operator, "*"], [:array, [:row, [:cell, "B1"]], [:row, [:cell, "B2"]]]]\n|)
    output = StringIO.new
    RewriteArrayFormulaeToArrays.rewrite(input,output)
    expected_output = <<END
B6\tB6:B8\t[:array, [:row, [:arithmetic, [:cell, "A1"], [:operator, "*"], [:cell, "B1"]]], [:row, [:arithmetic, [:cell, "A2"], [:operator, "*"], [:cell, "B2"]]]]
END
    output.string.should == expected_output
  end

end
