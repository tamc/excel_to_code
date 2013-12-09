require_relative '../../spec_helper'

describe CompileToRuby do

  
  def compile(input, sheet_names = {}, settable = nil)    
    output = StringIO.new
    c = CompileToRuby.new
    c.settable = settable
    c.rewrite(input,sheet_names,output)
    output.string
  end
  
it "should compile simple arithmetic" do
  input = {
    ["sheet1", "A1"] => [:arithmetic, [:number, "1"], [:operator, :+], [:number, "1"]],
    ["sheet1", "A2"] => [:arithmetic, [:number, "1"], [:operator, :-], [:number, "1"]],
    ["sheet1", "A3"] => [:arithmetic, [:number, "1"], [:operator, :*], [:number, "1"]],
    ["sheet1", "A4"] => [:arithmetic, [:number, "1"], [:operator, :/], [:number, "1"]],
    ["sheet1", "A5"] => [:arithmetic, [:number, "1"], [:operator, :^], [:number, "1"]],
    ["sheet1", "A6"] => [:arithmetic, [:number, "1.1"], [:operator, :+], [:number, "-1E12"]]
  }

expected = <<END
  def a1; @a1 ||= add(1,1); end
  def a2; @a2 ||= subtract(1,1); end
  def a3; @a3 ||= multiply(1,1); end
  def a4; @a4 ||= divide(1,1); end
  def a5; @a5 ||= power(1,1); end
  def a6; @a6 ||= add(1.1,-1000000000000.0); end
END

compile(input).should == expected
end

it "should compile references, mapping sheet references appropriately" do
  input = {
    ["sheet1", "A1"] => [:sheet_reference, "A complicated sheet name",[:cell, "A2"]]
  }

  sheet_names = {
    "A complicated sheet name" => "a_complicated_sheet_name"
  }

expected = <<END
  def a1; @a1 ||= a_complicated_sheet_name_a2; end
END

compile(input,sheet_names).should == expected
end

it "should compile references that are 'settable' as accessors" do
  input = {
    ["sheet1", "A1"] => [:number,1],
    ["sheet1", "A2"] => [:number,2],
    ["sheet1", "A3"] => [:number,3]
  }

sheet_names = {}

settable = lambda { |reference| reference == ["sheet1", "A2"] }

expected = <<END
  def a1; @a1 ||= 1; end
  attr_accessor :a2 # Default: 2
  def a3; @a3 ||= 3; end
END

compile(input,sheet_names,settable).should == expected
end

it "If has 'settable' accessors, and given a defaults file, should dump the default values to that file" do
input = {
["sheet1", "A1"] => [:number,1],
["sheet1", "A2"] => [:number,2],
["sheet1", "A3"] => [:number,3]
}

sheet_names = {}

settable = lambda { |reference| reference == ["sheet1", "A2"] }

expected_main = <<END
  def a1; @a1 ||= 1; end
  attr_accessor :a2 # Default: 2
  def a3; @a3 ||= 3; end
END

expected_defaults = [ "    @a2 = 2" ]

o = StringIO.new
c = CompileToRuby.new
c.settable = settable
c.rewrite(input, sheet_names, o)
o.string.should == expected_main
c.defaults.should == expected_defaults
end

end



