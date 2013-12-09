require_relative '../spec_helper'

describe ExtractNamedReferences do
  
  it "should output the named references in the workbook in the form <worksheet>\\t<name>\\t<reference>" do
    input = excel_fragment 'Workbook.xml'
    output = ExtractNamedReferences.extract(input)
    output.should == {
      "in_result" =>  "Inputs!$A$3",
      [:inputs, "local_named_reference"] => "Inputs!$A$3"
    }
  end
end
