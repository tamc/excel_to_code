require_relative '../spec_helper'

describe ReplaceRangesWithArrayLiterals do
  
  it "should replace ranges (e.g., A1:B2) with array literals (e.g., {A1,B1;A2,B2}) except if they only refer to a single cell, in which case replace with a single cell" do

input = <<END
A1\t[:sheet_reference, :sheet1, [:area, :A1, :B2]]
A2\t[:area, :"F$197", :"F$199"]
A3\t[:area, :"$F197", :"F$197"]
END

expected_output = <<END
A1\t[:array, [:row, [:sheet_reference, :sheet1, [:cell, :A1]], [:sheet_reference, :sheet1, [:cell, :B1]]], [:row, [:sheet_reference, :sheet1, [:cell, :A2]], [:sheet_reference, :sheet1, [:cell, :B2]]]]
A2\t[:array, [:row, [:cell, :F197]], [:row, [:cell, :F198]], [:row, [:cell, :F199]]]
A3\t[:cell, :F197]
END
    
    input = StringIO.new(input)
    output = StringIO.new
    ReplaceRangesWithArrayLiterals.replace(input,output)
    output.string.should == expected_output
  end

  # FIXME: Not working at the moment
  # it "should return exactly the same array literal for every sheet reference expansion" do
  #   testA = [:sheet_reference, :sheet1, [:area, :A1, :B2]]
  #   testB = [:sheet_reference, :sheet1, [:area, :A1, :B2]]
  #   r = ReplaceRangesWithArrayLiteralsAst.new
  #   first = r.map(testA)
  #   second = r.map(testB)
  #   first.object_id.should == second.object_id
  # end
  #
it "should deal with the edge case where SUMIF(A1:A10, 10, B5:B6) is interpreted by Excel as SUMIF(A1:A10, 10, B5:B15)" do

  r = ReplaceRangesWithArrayLiteralsAst.new

  ast = [:function, :SUMIF, [:sheet_reference, :sheet1, [:area, :A1, :A3]], [:number, 10], [:sheet_reference, :sheet1, [:area, :B1, :B2]]]
  r.map(ast).should == [:function, :SUMIF, 
                        [:array, 
                         [:row, [:sheet_reference, :sheet1, [:cell, :A1]]], 
                         [:row, [:sheet_reference, :sheet1, [:cell, :A2]]],
                         [:row, [:sheet_reference, :sheet1, [:cell, :A3]]]
                        ], [:number, 10], 
                        [:array, 
                         [:row, [:sheet_reference, :sheet1, [:cell, :B1]]], 
                         [:row, [:sheet_reference, :sheet1, [:cell, :B2]]],
                         [:row, [:sheet_reference, :sheet1, [:cell, :B3]]]
                        ]]
  ast = [:function, :SUMIF, [:sheet_reference, :sheet1, [:area, :A1, :A3]], [:number, 10], [:sheet_reference, :sheet1, [:cell, :B10]]]
  r.map(ast).should == [:function, :SUMIF, 
                        [:array, 
                         [:row, [:sheet_reference, :sheet1, [:cell, :A1]]], 
                         [:row, [:sheet_reference, :sheet1, [:cell, :A2]]],
                         [:row, [:sheet_reference, :sheet1, [:cell, :A3]]]
                        ], [:number, 10], 
                        [:array, 
                         [:row, [:sheet_reference, :sheet1, [:cell, :B10]]], 
                         [:row, [:sheet_reference, :sheet1, [:cell, :B11]]],
                         [:row, [:sheet_reference, :sheet1, [:cell, :B12]]]
                        ]]
  

end

end
