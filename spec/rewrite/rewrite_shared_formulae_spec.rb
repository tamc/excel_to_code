require_relative '../spec_helper'

describe RewriteSharedFormulae do
  
it "should take shared formula like 'reference\\tsharing range\\tsharing target identifier\\tformula ast\\n' and output a series of 'reference\\tformula ast\\n' for the shared formulae" do
input = StringIO.new("B3\tB3:C4\t0\t[:cell,'A1']\nZ5\tZ5\t1\t[:cell,'A1']\n")
shared_targets = StringIO.new("B3\t0\nB4\t0\nC3\t0\nC4\t0\nZ5\t1\n")
output = StringIO.new
RewriteSharedFormulae.rewrite(input,shared_targets,output)
expected_output = <<END
B3	[:cell, "A1"]
B4	[:cell, "A2"]
C3	[:cell, "B1"]
C4	[:cell, "B2"]
Z5	[:cell, "A1"]
END
output.string.should == expected_output
end

it "should take cope with occasions when the shared formula is not the top left in the range" do
input = StringIO.new("C3\tB3:C4\t0\t[:cell,'B1']\n")
shared_targets = StringIO.new("B3\t0\nB4\t0\nC3\t0\nC4\t0\nZ5\t0\n")
output = StringIO.new
RewriteSharedFormulae.rewrite(input,shared_targets,output)
expected_output = <<END
B3	[:cell, "A1"]
B4	[:cell, "A2"]
C3	[:cell, "B1"]
C4	[:cell, "B2"]
END
output.string.should == expected_output
end

it "should cope with exceptions in the shared formula range" do
input = StringIO.new("B3\tB3:C4\t0\t[:cell,'A1']\nZ5\tZ5\t1\t[:cell,'A1']\n")
shared_targets = StringIO.new("B3\t0\nC3\t0\nC4\t0\nZ5\t1\n")
output = StringIO.new
RewriteSharedFormulae.rewrite(input,shared_targets,output)
expected_output = <<END
B3	[:cell, "A1"]
C3	[:cell, "B1"]
C4	[:cell, "B2"]
Z5	[:cell, "A1"]
END
output.string.should == expected_output
end

it "should cope with overlapping shared formula ranges" do
input = StringIO.new("B3\tB3:C4\t0\t[:cell,'A1']\nB4\tB3:C4\t1\t[:cell,'C2']\n")
shared_targets = StringIO.new("B3\t0\nB4\t1\nC3\t1\nC4\t0\n")
output = StringIO.new
RewriteSharedFormulae.rewrite(input,shared_targets,output)
expected_output = <<END
B3	[:cell, "A1"]
C4	[:cell, "B2"]
B4	[:cell, "C2"]
C3	[:cell, "D1"]
END
output.string.should == expected_output
end

end
