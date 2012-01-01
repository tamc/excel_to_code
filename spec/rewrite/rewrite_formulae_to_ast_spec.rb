require_relative '../spec_helper'

describe RewriteFormulaeToAst do
  
  it "should take input in the form: 'thing\\tthing\\tformula\\n' where the last field is always a forumla and map the formulae to ast to produce 'thing\\tthing\\tast\\n'" do
    input = StringIO.new("B1\t1+1\nB3\tB3:B4\tCOSH(2*PI())\n")
    output = StringIO.new
    RewriteFormulaeToAst.rewrite(input,output)
    expected_output = <<END
B1\t[:formula, [:arithmetic, [:number, "1"], [:operator, "+"], [:number, "1"]]]
B3\tB3:B4\t[:formula, [:function, "COSH", [:arithmetic, [:number, "2"], [:operator, "*"], [:function, "PI"]]]]
END
    output.string.should == expected_output
  end
end
