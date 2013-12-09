require_relative '../../spec_helper'

describe CheckForUnknownFunctions do

it "should print a list of formulae that are in the input but have not been implemented by the compiler" do
input = <<END
A1\t[:function, :AVERAGE, [:number, 1]]
A2\t[:function, :"NOT IMPLEMENTED FORMULA", [:number, 1]]
END

expected = <<END
NOT IMPLEMENTED FORMULA
END

i = StringIO.new(input)
o = StringIO.new

c = CheckForUnknownFunctions.new
c.check(i,o)
o.string.should == expected
end
end



