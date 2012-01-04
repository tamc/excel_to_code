require_relative '../spec_helper'

describe RewriteSharedFormulae do
  
  it "should take shared formula like 'reference\\tsharing range\\tformula ast\\n' and output a series of 'reference\\tformula ast\\n' for the shared formulae" do
    input = StringIO.new("B3\tB3:C4\t[:formula,[:cell,'A1']]\n")
    output = StringIO.new
    RewriteSharedFormulae.rewrite(input,output)
    expected_output = <<END
B3	[:formula, [:cell, "A1"]]
B4	[:formula, [:cell, "A2"]]
C3	[:formula, [:cell, "B1"]]
C4	[:formula, [:cell, "B2"]]
END
    output.string.should == expected_output
  end
end
