require_relative '../../spec_helper'

describe CompileToRuby do

  
  def compile(text, sheet_names = "", settable = nil)    
    input = StringIO.new(text)
    sheet_names = StringIO.new(sheet_names)
    output = StringIO.new
    c = CompileToRuby.new
    c.settable = settable
    c.rewrite(input,sheet_names,output)
    output.string
  end
  
it "should compile simple arithmetic" do
input = <<END
A1\t[:arithmetic, [:number, "1"], [:operator, "+"], [:number, "1"]]
A2\t[:arithmetic, [:number, "1"], [:operator, "-"], [:number, "1"]]
A3\t[:arithmetic, [:number, "1"], [:operator, "*"], [:number, "1"]]
A4\t[:arithmetic, [:number, "1"], [:operator, "/"], [:number, "1"]]
A5\t[:arithmetic, [:number, "1"], [:operator, "^"], [:number, "1"]]
A6\t[:arithmetic, [:number, "1.1"], [:operator, "+"], [:number, "-1E12"]]
END

expected = <<END
  def a1; add(1,1); end
  def a2; subtract(1,1); end
  def a3; multiply(1,1); end
  def a4; divide(1,1); end
  def a5; power(1,1); end
  def a6; add(1.1,-1000000000000.0); end
END

compile(input).should == expected
end

it "should compile references, mapping sheet references appropriately" do
input = <<END
A1\t[:sheet_reference, "A complicated sheet name",[:cell, "A2"]]
END

sheet_names = <<END
A complicated sheet name\ta_complicated_sheet_name
END

expected = <<END
  def a1; a_complicated_sheet_name.a2; end
END

compile(input,sheet_names).should == expected
end

it "should compile references that are 'settable' as accessors" do
input = <<END
A1\t[:number,1]
A2\t[:number,2]
A3\t[:number,3]
END

sheet_names = <<END
END

settable = lambda do |reference|
  if reference == 'A2'
    true
  else
    false
  end
end

expected = <<END
  def a1; 1; end
  attr_accessor :a2 # Default: 2
  def a3; 3; end
END

compile(input,sheet_names,settable).should == expected
end

it "If has 'settable' accessors, and given a defaults file, should dump the default values to that file" do
input = <<END
A1\t[:number,1]
A2\t[:number,2]
A3\t[:number,3]
END

sheet_names = <<END
END

settable = lambda do |reference|
  if reference == 'A2'
    true
  else
    false
  end
end

expected_main = <<END
  def a1; 1; end
  attr_accessor :a2 # Default: 2
  def a3; 3; end
END

expected_defaults = <<END
    @a2 = 2
END

i = StringIO.new(input)
s = StringIO.new(sheet_names)
o = StringIO.new
d = StringIO.new
c = CompileToRuby.new
c.settable = settable
c.rewrite(i,s,o,d)

o.string.should == expected_main
d.string.should == expected_defaults
end

end



