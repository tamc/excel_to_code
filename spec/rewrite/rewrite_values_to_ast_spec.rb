require_relative '../spec_helper'

describe RewriteValuesToAst do
  
  it "should create a flat file with one string per cell, in the format: reference\ttype\tvalue" do
    input = <<END
A1\tb\t1
A2\ts\t0
A3\tn\t1
A4\tn\t3.1415000000000002
A5\te\t#NAME?
A6\tstr\tHello    
AJ5	str	Capacity increases up to ~8,5 GW in 2050._x000D_This requires installing on average 300 MW, or ~120 new turbines per year
END
    output = StringIO.new
    RewriteValuesToAst.rewrite(input,output)
    expected_output = <<END
A1\t[:boolean_true]
A2\t[:shared_string, "0"]
A3\t[:number, "1"]
A4\t[:number, "3.1415000000000002"]
A5\t[:error, "#NAME?"]
A6\t[:string, "Hello    "]
AJ5\t[:string, "Capacity increases up to ~8,5 GW in 2050.This requires installing on average 300 MW, or ~120 new turbines per year"]
END
    output.string.should == expected_output
  end
end
