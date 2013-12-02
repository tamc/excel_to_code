require_relative '../spec_helper'

describe ExtractSharedFormulaeTargets do
  
  it "should create a hash with ['sheetname', 'B1'] => '0' where '0' is the shared_target_number " do
    input = excel_fragment 'FormulaeTypes.xml'
    output = ExtractSharedFormulaeTargets.extract('SheetName', input)
    output.should == {
      ["SheetName", "B3"] => "0",
      ["SheetName", "B4"] => "0",
    }
  end
end
