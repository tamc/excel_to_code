require_relative '../spec_helper'

describe RemoveCells do

  it "should remove any cells whose references are NOT in the given array" do
    input = {
      ['sheet1', 'A1'] => [:boolean_true],
      ['sheet1', 'A2'] => [:shared_string, "0"],
      ['sheet1', 'A3'] => [:number, "1"],
      ['sheet1', 'A4'] => [:number, "3.1415000000000002"],
      ['sheet1', 'A5'] => [:error, "#NAME?"],
      ['sheet1', 'A6'] => [:string, "Hello    "],
    }

    cells_to_keep = {'sheet1' => {'A1' => true,'A2' => true, 'A6' => true}}

    expected_output = {
      ['sheet1', 'A1'] => [:boolean_true],
      ['sheet1', 'A2'] => [:shared_string, "0"],
      ['sheet1', 'A6'] => [:string, "Hello    "],
    }

    r = RemoveCells.new
    r.cells_to_keep = cells_to_keep
    r.rewrite(input).should == expected_output

  end # / it

end # /describe
