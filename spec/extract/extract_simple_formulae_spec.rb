require_relative '../spec_helper'

describe ExtractSimpleFormulae do
  
  it "should create a hash ['SheetName', 'A1'] => '1+1' ... " do
    input = excel_fragment 'FormulaeTypes.xml'
    output = ExtractSimpleFormulae.extract("SheetName", input)
    output.should == {
      [:SheetName, :B1] => [:arithmetic, [:number, 1.0], [:operator, "+"], [:number, 1.0]],
      [:SheetName, :B2] => [:function, "COSH", [:arithmetic, [:number, 2.0], [:operator, "*"], [:function, "PI"]]]
    }
  end
end
