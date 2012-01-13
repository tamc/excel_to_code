require_relative '../spec_helper'

describe RewriteArrayFormulae do
  
  it "should take array formula like 'reference\\array range\\tformula ast\\n' and output normal formulae" do
    input = StringIO.new(%Q|B6\tB6:B8\t[:arithmetic, [:area, "B3", "B5"], [:operator, "+"], [:area, "C3", "C5"]]\n|)
    output = StringIO.new
    RewriteArrayFormulae.rewrite(input,output)
    expected_output = <<END
B6	[:arithmetic, [:cell, "B3"], [:operator, "+"], [:cell, "C3"]]
B7	[:arithmetic, [:cell, "B4"], [:operator, "+"], [:cell, "C4"]]
B8	[:arithmetic, [:cell, "B5"], [:operator, "+"], [:cell, "C5"]]
END
    output.string.should == expected_output
  end

  it "should work with single cell array formulae" do
    input = StringIO.new(%Q|B6\tB6\t[:arithmetic, [:area, "B3", "B5"], [:operator, "+"], [:area, "C3", "C5"]]\n|)
    output = StringIO.new
    RewriteArrayFormulae.rewrite(input,output)
    expected_output = <<END
B6	[:arithmetic, [:cell, "B3"], [:operator, "+"], [:cell, "C3"]]
END
    output.string.should == expected_output
  end


end
