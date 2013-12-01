require_relative '../spec_helper'

describe ExtractNamedReferences do
  
  it "should output the named references in the workbook in the form <worksheet>\\t<name>\\t<reference>" do
    input = excel_fragment 'Workbook.xml'
    output = ExtractNamedReferences.extract(input)
    output.should == {
      "In_result" =>  "Inputs!$A$3",
      ["Inputs", "Local_named_reference"] => "Inputs!$A$3"
    }
  end
end
