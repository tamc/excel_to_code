require_relative '../spec_helper'

describe ExtractValues do
  
  it "should create a hash with the types and values of the cells" do
    input = excel_fragment 'ValueTypes.xml'
    output = ExtractValues.extract("sheet_name", input)
    expected_output = {
      ["sheet_name", "A1"]  => [:boolean_true],
      ["sheet_name", "A2"]  => [:shared_string, "0"],
      ["sheet_name", "A3"]  => [:number, "1"],
      ["sheet_name", "A4"]  => [:number, "3.1415000000000002"],
      ["sheet_name", "A5"]  => [:error, "#NAME?"],
      ["sheet_name", "A6"]  => [:string, "Hello"],
    }
    output.should == expected_output
  end
end
