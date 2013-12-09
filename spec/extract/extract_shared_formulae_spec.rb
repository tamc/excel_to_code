require_relative '../spec_helper'

describe ExtractSharedFormulae do
  
  it "should create a hash of SharedFormulae instances" do
    input = excel_fragment 'FormulaeTypes.xml'
    output = ExtractSharedFormulae.extract("SheetName", input)
    output.should == {
      [:SheetName, :B3] => ["B3:B4", "0", [:function, :COSH, [:arithmetic, [:number, 2.0], [:operator, :*], [:function, :PI]]]]
    }
  end
end
