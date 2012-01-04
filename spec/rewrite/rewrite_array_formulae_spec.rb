require_relative '../spec_helper'

describe RewriteArrayFormulae do
  
  it "should take array formula like 'reference\\array range\\tformula ast\\n' and output a series of 'reference\\tformula ast\\n' for the array formulae" do
    input = StringIO.new(%Q|B6\tB6:B8\t[:formula, [:arithmetic, [:area, "B3", "B5"], [:operator, "+"], [:area, "C3", "C5"]]]\n|)
    output = StringIO.new
    RewriteArrayFormulae.rewrite(input,output)
    expected_output = <<END
B6\t[:formula, [:function, "ARRAYFORMULA", [:arithmetic, [:area, "B3", "B5"], [:operator, "+"], [:area, "C3", "C5"]]]]
B7\t[:formula, [:function, "CONTINUE", [:cell, "B6"], 2, 1]]
B8\t[:formula, [:function, "CONTINUE", [:cell, "B6"], 3, 1]]
END
    output.string.should == expected_output
  end
end
