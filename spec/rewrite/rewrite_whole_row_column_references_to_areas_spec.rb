require_relative '../spec_helper'

describe RewriteWholeRowColumnReferencesToAreas do
  
  it "should take a file with ast in its last column, and a file with worksheetname\\tdimensionrange\\n and map any row and column references in the ast to conventional area references" do
    input = test_data('RowAndColumnRanges.ast')
    default_worksheet_name = "Ranges"
    worksheet_dimensions = test_data("RowAndColumnRangeDimensions")
    output = StringIO.new
    RewriteWholeRowColumnReferencesToAreas.rewrite(input,default_worksheet_name,worksheet_dimensions,output)
    expected =<<END
B2	[:formula, [:function, "SUM", [:area, "F4", "F6"]]]
C2	[:formula, [:function, "SUM", [:sheet_reference, "ValueTypes", [:area, "A3", "A4"]]]]
B3	[:formula, [:function, "SUM", [:area, "F1", "F6"]]]
C3	[:formula, [:function, "SUM", [:sheet_reference, "ValueTypes", [:area, "A1", "A6"]]]]
B4	[:formula, [:function, "SUM", [:area, "A5", "G5"]]]
C4	[:formula, [:function, "SUM", [:sheet_reference, "ValueTypes", [:area, "A4", "A4"]]]]
END
    output.string.should == expected
  end
end
