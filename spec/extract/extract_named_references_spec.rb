require_relative '../spec_helper'

describe ExtractNamedReferences do
  
  it "should output the named references in the workbook in the form <worksheet>\\t<name>\\t<reference>" do
    input = excel_fragment 'Workbook.xml'
    output = StringIO.new
    ExtractNamedReferences.extract(input,output)
    output.string.should == <<END
\tIn_result\tInputs!$A$3
Inputs\tLocal_named_reference\tInputs!$A$3
END
  end
end
