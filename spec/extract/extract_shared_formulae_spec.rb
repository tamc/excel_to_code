require_relative '../spec_helper'

describe ExtractSharedFormulae do
  
  it "should create a flat file with one string per formula, in the format: reference\tsharing range\tshared_formula_identifier\tformula" do
    input = excel_fragment 'FormulaeTypes.xml'
    output = StringIO.new
    ExtractSharedFormulae.extract(input,output)
    expected_output = <<END
B3\tB3:B4\t0\tCOSH(2*PI())
END
    output.string.should == expected_output
  end
end
