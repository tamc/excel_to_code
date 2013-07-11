require_relative '../spec_helper'

describe ExtractSharedFormulaeTargets do
  
  it "should create a flat file with one line per formula if it is the target of a shared formula in the format: reference\tshared_group_number\n" do
    input = excel_fragment 'FormulaeTypes.xml'
    output = StringIO.new
    ExtractSharedFormulaeTargets.extract(input,output)
    expected_output = <<END
B3\t0
B4\t0
END
    output.string.should == expected_output
  end
end
