require_relative '../spec_helper'
require_relative '../../src/extract/extract_array_formulae'
require 'stringio'

describe ExtractArrayFormulae do
  
  it "should create a flat file with one string per formula, in the format: reference\tarray range\tformula" do
    input = excel_fragment 'FormulaeTypes.xml'
    output = StringIO.new
    ExtractArrayFormulae.extract(input,output)
    expected_output = <<END
B5\tB5\tB1:B4
B6\tB6:B8\tIF(B3:B5=8,"Eight","Not Eight")
END
    output.string.should == expected_output
  end
end
