require_relative '../spec_helper'
require_relative '../../src/extract/extract_values'
require 'stringio'

describe ExtractValues do
  
  it "should create a flat file with one string per cell, in the format: reference\ttype\tvalue" do
    input = excel_fragment 'ValueTypes.xml'
    output = StringIO.new
    ExtractValues.extract(input,output)
    expected_output = <<END
A1\tb\t1
A2\ts\t0
A3\tn\t1
A4\tn\t3.1415000000000002
A5\te\t#NAME?
A6\tstr\tHello
END
    output.string.should == expected_output
  end
end
