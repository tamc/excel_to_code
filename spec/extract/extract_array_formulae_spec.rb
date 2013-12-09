require_relative '../spec_helper'

describe ExtractArrayFormulae do
  
  it "should create a Hash like ['SheetName', 'B3'] => ['B6:B8', [ast..]]" do
    input = excel_fragment 'FormulaeTypes.xml'
    output = ExtractArrayFormulae.extract("SheetName", input)
    output.should == {
      [:SheetName, :B5] => ["B5", [:area, :B1, :B4]],
      [:SheetName, :B6] => ["B6:B8", [:function, :IF, [:comparison, [:area, :B3, :B5], [:comparator, :"="], [:number, 8.0]], [:string, "Eight"], [:string, "Not Eight"]]],
    }
  end
end
