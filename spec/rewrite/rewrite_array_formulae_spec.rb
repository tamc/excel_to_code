require_relative '../spec_helper'

describe RewriteArrayFormulae do
  
  it "should take array formula like 'reference\\array range\\tformula ast\\n' and output normal formulae" do
    input = StringIO.new(%Q|B6\tB6:B8\t[:array, [:row, [:arithmetic, [:cell, "B3"], [:operator, "+"], [:cell, "C3"]]], [:row, [:arithmetic, [:cell, "B4"], [:operator, "+"], [:cell, "C4"]]], [:row, [:arithmetic, [:cell, "B5"], [:operator, "+"], [:cell, "C5"]]]]\n|)
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
    input = StringIO.new(%Q|B6\tB6\t[:array, [:row, [:arithmetic, [:cell, "B3"], [:operator, "+"], [:cell, "C3"]]], [:row, [:arithmetic, [:cell, "B4"], [:operator, "+"], [:cell, "C4"]]], [:row, [:arithmetic, [:cell, "B5"], [:operator, "+"], [:cell, "C5"]]]]\n|)
    output = StringIO.new
    RewriteArrayFormulae.rewrite(input,output)
    expected_output = <<END
B6	[:arithmetic, [:cell, "B3"], [:operator, "+"], [:cell, "C3"]]
END
    output.string.should == expected_output
  end
  
it "if the array is a single column, should repeat the column across columns in the output" do
    input = StringIO.new(%Q|B6\tB6:C9\t[:array, [:row, [:arithmetic, [:cell, "B3"], [:operator, "+"], [:cell, "C3"]]], [:row, [:arithmetic, [:cell, "B4"], [:operator, "+"], [:cell, "C4"]]], [:row, [:arithmetic, [:cell, "B5"], [:operator, "+"], [:cell, "C5"]]]]\n|)
    output = StringIO.new
    RewriteArrayFormulae.rewrite(input,output)
    expected_output = <<END
B6	[:arithmetic, [:cell, "B3"], [:operator, "+"], [:cell, "C3"]]
B7	[:arithmetic, [:cell, "B4"], [:operator, "+"], [:cell, "C4"]]
B8	[:arithmetic, [:cell, "B5"], [:operator, "+"], [:cell, "C5"]]
B9	[:error, "#N/A"]
C6	[:arithmetic, [:cell, "B3"], [:operator, "+"], [:cell, "C3"]]
C7	[:arithmetic, [:cell, "B4"], [:operator, "+"], [:cell, "C4"]]
C8	[:arithmetic, [:cell, "B5"], [:operator, "+"], [:cell, "C5"]]
C9	[:error, "#N/A"]
END
    output.string.should == expected_output
  end
  
it "if the array is a single row, should repeat the row across rows in the output" do
    input = StringIO.new(%Q|B6\tB6:D9\t[:array, [:row, [:arithmetic, [:cell, "B3"], [:operator, "+"], [:cell, "C3"]], [:arithmetic, [:cell, "B4"], [:operator, "+"], [:cell, "C4"]]]]\n|)
    output = StringIO.new
    RewriteArrayFormulae.rewrite(input,output)
    expected_output = <<END
B6	[:arithmetic, [:cell, "B3"], [:operator, "+"], [:cell, "C3"]]
B7	[:arithmetic, [:cell, "B3"], [:operator, "+"], [:cell, "C3"]]
B8	[:arithmetic, [:cell, "B3"], [:operator, "+"], [:cell, "C3"]]
B9	[:arithmetic, [:cell, "B3"], [:operator, "+"], [:cell, "C3"]]
C6	[:arithmetic, [:cell, "B4"], [:operator, "+"], [:cell, "C4"]]
C7	[:arithmetic, [:cell, "B4"], [:operator, "+"], [:cell, "C4"]]
C8	[:arithmetic, [:cell, "B4"], [:operator, "+"], [:cell, "C4"]]
C9	[:arithmetic, [:cell, "B4"], [:operator, "+"], [:cell, "C4"]]
D6	[:error, "#N/A"]
D7	[:error, "#N/A"]
D8	[:error, "#N/A"]
D9	[:error, "#N/A"]
END
    output.string.should == expected_output
end

it "if the array is a single cell, should repeat the row across rows and columns in the output" do
    input = StringIO.new(%Q|B6\tB6:D9\t[:array, [:row, [:arithmetic, [:cell, "B3"], [:operator, "+"], [:cell, "C3"]]]]\n|)
    output = StringIO.new
    RewriteArrayFormulae.rewrite(input,output)
    expected_output = <<END
B6	[:arithmetic, [:cell, "B3"], [:operator, "+"], [:cell, "C3"]]
B7	[:arithmetic, [:cell, "B3"], [:operator, "+"], [:cell, "C3"]]
B8	[:arithmetic, [:cell, "B3"], [:operator, "+"], [:cell, "C3"]]
B9	[:arithmetic, [:cell, "B3"], [:operator, "+"], [:cell, "C3"]]
C6	[:arithmetic, [:cell, "B3"], [:operator, "+"], [:cell, "C3"]]
C7	[:arithmetic, [:cell, "B3"], [:operator, "+"], [:cell, "C3"]]
C8	[:arithmetic, [:cell, "B3"], [:operator, "+"], [:cell, "C3"]]
C9	[:arithmetic, [:cell, "B3"], [:operator, "+"], [:cell, "C3"]]
D6	[:arithmetic, [:cell, "B3"], [:operator, "+"], [:cell, "C3"]]
D7	[:arithmetic, [:cell, "B3"], [:operator, "+"], [:cell, "C3"]]
D8	[:arithmetic, [:cell, "B3"], [:operator, "+"], [:cell, "C3"]]
D9	[:arithmetic, [:cell, "B3"], [:operator, "+"], [:cell, "C3"]]
END
    output.string.should == expected_output
end

  
  
it "should deal with repetition of array formula that only produce a single answer" do
    input = StringIO.new(%Q|B6\tB6:B8\t[:function, "SUM", [:array, [:row, [:arithmetic, [:cell, "B3"], [:operator, "+"], [:cell, "C3"]]], [:row, [:arithmetic, [:cell, "B4"], [:operator, "+"], [:cell, "C4"]]], [:row, [:arithmetic, [:cell, "B5"], [:operator, "+"], [:cell, "C5"]]]]]\n|)
    output = StringIO.new
    RewriteArrayFormulae.rewrite(input,output)
    expected_output = <<END
B6	[:function, "SUM", [:array, [:row, [:arithmetic, [:cell, "B3"], [:operator, "+"], [:cell, "C3"]]], [:row, [:arithmetic, [:cell, "B4"], [:operator, "+"], [:cell, "C4"]]], [:row, [:arithmetic, [:cell, "B5"], [:operator, "+"], [:cell, "C5"]]]]]
B7	[:function, "SUM", [:array, [:row, [:arithmetic, [:cell, "B3"], [:operator, "+"], [:cell, "C3"]]], [:row, [:arithmetic, [:cell, "B4"], [:operator, "+"], [:cell, "C4"]]], [:row, [:arithmetic, [:cell, "B5"], [:operator, "+"], [:cell, "C5"]]]]]
B8	[:function, "SUM", [:array, [:row, [:arithmetic, [:cell, "B3"], [:operator, "+"], [:cell, "C3"]]], [:row, [:arithmetic, [:cell, "B4"], [:operator, "+"], [:cell, "C4"]]], [:row, [:arithmetic, [:cell, "B5"], [:operator, "+"], [:cell, "C5"]]]]]
END
    output.string.should == expected_output
  end


end
