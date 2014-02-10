require_relative '../spec_helper'

describe ReplaceArithmeticOnRanges do
  
  it "should replace arithmetic on ranges (e.g., 1/A1:B2) with individual calculations (i.e., [1/A1, 1/B1, 1/A2, 1/B2])" do

input = <<END
A0\t[:arithmetic, [:number, "1"], [:operator, "/"], [cell, "F198]]
A1\t[:arithmetic, [:number, "1"], [:operator, "/"], [:array, [:row, [:cell, "F197"]], [:row, [:cell, "F198"]], [:row, [:cell, "F199"]]]]
A2\t[:arithmetic, [:array, [:row, [:cell, "F197"]], [:row, [:cell, "F198"]], [:row, [:cell, "F199"]]], [:operator, "+"], [:number, "1"]]
A3\t[:arithmetic, [:array, [:row, [:cell, "F197"]], [:row, [:cell, "F198"]], [:row, [:cell, "F199"]]], [:operator, "/"],  [:array, [:row, [:cell, "F197"]], [:row, [:cell, "F198"]], [:row, [:cell, "F199"]]]]
END

expected_output = <<END
A0\t[:arithmetic, [:number, "1"], [:operator, "/"], [cell, "F198]]
A1\t[:array, [:row, [:arithmetic, [:number, "1"], [:operator, "/"], [:cell, "F197"]]], [:row, [:arithmetic, [:number, "1"], [:operator, "/"], [:cell, "F198"]]], [:row, [:arithmetic, [:number, "1"], [:operator, "/"], [:cell, "F199"]]]]
A2\t[:array, [:row, [:arithmetic, [:cell, "F197"], [:operator, "+"], [:number, "1"]]], [:row, [:arithmetic, [:cell, "F198"], [:operator, "+"], [:number, "1"]]], [:row, [:arithmetic, [:cell, "F199"], [:operator, "+"], [:number, "1"]]]]
A3\t[:array, [:row, [:arithmetic, [:cell, "F197"], [:operator, "/"], [:cell, "F197"]]], [:row, [:arithmetic, [:cell, "F198"], [:operator, "/"], [:cell, "F198"]]], [:row, [:arithmetic, [:cell, "F199"], [:operator, "/"], [:cell, "F199"]]]]
END
    
    input = StringIO.new(input)
    output = StringIO.new
    ReplaceArithmeticOnRanges.replace(input,output)
    output.string.should == expected_output
  end

  it "should work in complex nested cases" do
    input_ast = [:function,
                 :SUM,
                 [:arithmetic,
                  [:arithmetic,
                   [:array,
                    [:row,
                     [:sheet_reference, :"COM.DMD", [:cell, :K28]],
                     [:sheet_reference, :"COM.DMD", [:cell, :L28]],
                     [:sheet_reference, :"COM.DMD", [:cell, :M28]],
                     [:sheet_reference, :"COM.DMD", [:cell, :N28]],
                     [:sheet_reference, :"COM.DMD", [:cell, :O28]],
                     [:sheet_reference, :"COM.DMD", [:cell, :P28]],
                     [:sheet_reference, :"COM.DMD", [:cell, :Q28]],
                     [:sheet_reference, :"COM.DMD", [:cell, :R28]],
                     [:sheet_reference, :"COM.DMD", [:cell, :S28]],
                     [:sheet_reference, :"COM.DMD", [:cell, :T28]],
                     [:sheet_reference, :"COM.DMD", [:cell, :U28]],
                     [:sheet_reference, :"COM.DMD", [:cell, :V28]],
                     [:sheet_reference, :"COM.DMD", [:cell, :W28]],
                     [:sheet_reference, :"COM.DMD", [:cell, :X28]],
                     [:sheet_reference, :"COM.DMD", [:cell, :Y28]],
                     [:sheet_reference, :"COM.DMD", [:cell, :Z28]],
                     [:sheet_reference, :"COM.DMD", [:cell, :AA28]],
                     [:sheet_reference, :"COM.DMD", [:cell, :AB28]],
                     [:sheet_reference, :"COM.DMD", [:cell, :AC28]],
                     [:sheet_reference, :"COM.DMD", [:cell, :AD28]],
                     [:sheet_reference, :"COM.DMD", [:cell, :AE28]],
                     [:sheet_reference, :"COM.DMD", [:cell, :AF28]],
                     [:sheet_reference, :"COM.DMD", [:cell, :AG28]],
                     [:sheet_reference, :"COM.DMD", [:cell, :AH28]],
                     [:sheet_reference, :"COM.DMD", [:cell, :AI28]],
                     [:sheet_reference, :"COM.DMD", [:cell, :AJ28]],
                     [:sheet_reference, :"COM.DMD", [:cell, :AK28]],
                     [:sheet_reference, :"COM.DMD", [:cell, :AL28]],
                     [:sheet_reference, :"COM.DMD", [:cell, :AM28]],
                     [:sheet_reference, :"COM.DMD", [:cell, :AN28]],
                     [:sheet_reference, :"COM.DMD", [:cell, :AO28]]]],
                   [:operator, :*],
                   [:comparison,
                    [:sheet_reference, :"COM.DMD", [:cell, :L6]],
                    [:comparator, :"="],
                    [:array,
                     [:row,
                      [:sheet_reference, :"COM.DMD", [:cell, :K6]],
                      [:sheet_reference, :"COM.DMD", [:cell, :L6]],
                      [:sheet_reference, :"COM.DMD", [:cell, :M6]],
                      [:sheet_reference, :"COM.DMD", [:cell, :N6]],
                      [:sheet_reference, :"COM.DMD", [:cell, :O6]],
                      [:sheet_reference, :"COM.DMD", [:cell, :P6]],
                      [:sheet_reference, :"COM.DMD", [:cell, :Q6]],
                      [:sheet_reference, :"COM.DMD", [:cell, :R6]],
                      [:sheet_reference, :"COM.DMD", [:cell, :S6]],
                      [:sheet_reference, :"COM.DMD", [:cell, :T6]],
                      [:sheet_reference, :"COM.DMD", [:cell, :U6]],
                      [:sheet_reference, :"COM.DMD", [:cell, :V6]],
                      [:sheet_reference, :"COM.DMD", [:cell, :W6]],
                      [:sheet_reference, :"COM.DMD", [:cell, :X6]],
                      [:sheet_reference, :"COM.DMD", [:cell, :Y6]],
                      [:sheet_reference, :"COM.DMD", [:cell, :Z6]],
                      [:sheet_reference, :"COM.DMD", [:cell, :AA6]],
                      [:sheet_reference, :"COM.DMD", [:cell, :AB6]],
                      [:sheet_reference, :"COM.DMD", [:cell, :AC6]],
                      [:sheet_reference, :"COM.DMD", [:cell, :AD6]],
                      [:sheet_reference, :"COM.DMD", [:cell, :AE6]],
                      [:sheet_reference, :"COM.DMD", [:cell, :AF6]],
                      [:sheet_reference, :"COM.DMD", [:cell, :AG6]],
                      [:sheet_reference, :"COM.DMD", [:cell, :AH6]],
                      [:sheet_reference, :"COM.DMD", [:cell, :AI6]],
                      [:sheet_reference, :"COM.DMD", [:cell, :AJ6]],
                      [:sheet_reference, :"COM.DMD", [:cell, :AK6]],
                      [:sheet_reference, :"COM.DMD", [:cell, :AL6]],
                      [:sheet_reference, :"COM.DMD", [:cell, :AM6]],
                      [:sheet_reference, :"COM.DMD", [:cell, :AN6]],
                      [:sheet_reference, :"COM.DMD", [:cell, :AO6]]]]]],
                  [:operator, :*],
                  [:comparison,
                   [:sheet_reference, :"COM.DMD", [:cell, :B30]],
                   [:comparator, :"="],
                   [:array,
                    [:row,
                     [:sheet_reference, :"COM.DMD", [:cell, :K9]],
                     [:sheet_reference, :"COM.DMD", [:cell, :L9]],
                     [:sheet_reference, :"COM.DMD", [:cell, :M9]],
                     [:sheet_reference, :"COM.DMD", [:cell, :N9]],
                     [:sheet_reference, :"COM.DMD", [:cell, :O9]],
                     [:sheet_reference, :"COM.DMD", [:cell, :P9]],
                     [:sheet_reference, :"COM.DMD", [:cell, :Q9]],
                     [:sheet_reference, :"COM.DMD", [:cell, :R9]],
                     [:sheet_reference, :"COM.DMD", [:cell, :S9]],
                     [:sheet_reference, :"COM.DMD", [:cell, :T9]],
                     [:sheet_reference, :"COM.DMD", [:cell, :U9]],
                     [:sheet_reference, :"COM.DMD", [:cell, :V9]],
                     [:sheet_reference, :"COM.DMD", [:cell, :W9]],
                     [:sheet_reference, :"COM.DMD", [:cell, :X9]],
                     [:sheet_reference, :"COM.DMD", [:cell, :Y9]],
                     [:sheet_reference, :"COM.DMD", [:cell, :Z9]],
                     [:sheet_reference, :"COM.DMD", [:cell, :AA9]],
                     [:sheet_reference, :"COM.DMD", [:cell, :AB9]],
                     [:sheet_reference, :"COM.DMD", [:cell, :AC9]],
                     [:sheet_reference, :"COM.DMD", [:cell, :AD9]],
                     [:sheet_reference, :"COM.DMD", [:cell, :AE9]],
                     [:sheet_reference, :"COM.DMD", [:cell, :AF9]],
                     [:sheet_reference, :"COM.DMD", [:cell, :AG9]],
                     [:sheet_reference, :"COM.DMD", [:cell, :AH9]],
                     [:sheet_reference, :"COM.DMD", [:cell, :AI9]],
                     [:sheet_reference, :"COM.DMD", [:cell, :AJ9]],
                     [:sheet_reference, :"COM.DMD", [:cell, :AK9]],
                     [:sheet_reference, :"COM.DMD", [:cell, :AL9]],
                     [:sheet_reference, :"COM.DMD", [:cell, :AM9]],
                     [:sheet_reference, :"COM.DMD", [:cell, :AN9]],
                     [:sheet_reference, :"COM.DMD", [:cell, :AO9]]]]]]]

    output_ast = [:function, :SUM, [:array, [:row, [:arithmetic, [:arithmetic, [:sheet_reference, :"COM.DMD", [:cell, :K28]], [:operator, :*], [:comparison, [:sheet_reference, :"COM.DMD", [:cell, :L6]], [:comparator, :"="], [:sheet_reference, :"COM.DMD", [:cell, :K6]]]], [:operator, :*], [:comparison, [:sheet_reference, :"COM.DMD", [:cell, :B30]], [:comparator, :"="], [:sheet_reference, :"COM.DMD", [:cell, :K9]]]], [:arithmetic, [:arithmetic, [:sheet_reference, :"COM.DMD", [:cell, :L28]], [:operator, :*], [:comparison, [:sheet_reference, :"COM.DMD", [:cell, :L6]], [:comparator, :"="], [:sheet_reference, :"COM.DMD", [:cell, :L6]]]], [:operator, :*], [:comparison, [:sheet_reference, :"COM.DMD", [:cell, :B30]], [:comparator, :"="], [:sheet_reference, :"COM.DMD", [:cell, :L9]]]], [:arithmetic, [:arithmetic, [:sheet_reference, :"COM.DMD", [:cell, :M28]], [:operator, :*], [:comparison, [:sheet_reference, :"COM.DMD", [:cell, :L6]], [:comparator, :"="], [:sheet_reference, :"COM.DMD", [:cell, :M6]]]], [:operator, :*], [:comparison, [:sheet_reference, :"COM.DMD", [:cell, :B30]], [:comparator, :"="], [:sheet_reference, :"COM.DMD", [:cell, :M9]]]], [:arithmetic, [:arithmetic, [:sheet_reference, :"COM.DMD", [:cell, :N28]], [:operator, :*], [:comparison, [:sheet_reference, :"COM.DMD", [:cell, :L6]], [:comparator, :"="], [:sheet_reference, :"COM.DMD", [:cell, :N6]]]], [:operator, :*], [:comparison, [:sheet_reference, :"COM.DMD", [:cell, :B30]], [:comparator, :"="], [:sheet_reference, :"COM.DMD", [:cell, :N9]]]], [:arithmetic, [:arithmetic, [:sheet_reference, :"COM.DMD", [:cell, :O28]], [:operator, :*], [:comparison, [:sheet_reference, :"COM.DMD", [:cell, :L6]], [:comparator, :"="], [:sheet_reference, :"COM.DMD", [:cell, :O6]]]], [:operator, :*], [:comparison, [:sheet_reference, :"COM.DMD", [:cell, :B30]], [:comparator, :"="], [:sheet_reference, :"COM.DMD", [:cell, :O9]]]], [:arithmetic, [:arithmetic, [:sheet_reference, :"COM.DMD", [:cell, :P28]], [:operator, :*], [:comparison, [:sheet_reference, :"COM.DMD", [:cell, :L6]], [:comparator, :"="], [:sheet_reference, :"COM.DMD", [:cell, :P6]]]], [:operator, :*], [:comparison, [:sheet_reference, :"COM.DMD", [:cell, :B30]], [:comparator, :"="], [:sheet_reference, :"COM.DMD", [:cell, :P9]]]], [:arithmetic, [:arithmetic, [:sheet_reference, :"COM.DMD", [:cell, :Q28]], [:operator, :*], [:comparison, [:sheet_reference, :"COM.DMD", [:cell, :L6]], [:comparator, :"="], [:sheet_reference, :"COM.DMD", [:cell, :Q6]]]], [:operator, :*], [:comparison, [:sheet_reference, :"COM.DMD", [:cell, :B30]], [:comparator, :"="], [:sheet_reference, :"COM.DMD", [:cell, :Q9]]]], [:arithmetic, [:arithmetic, [:sheet_reference, :"COM.DMD", [:cell, :R28]], [:operator, :*], [:comparison, [:sheet_reference, :"COM.DMD", [:cell, :L6]], [:comparator, :"="], [:sheet_reference, :"COM.DMD", [:cell, :R6]]]], [:operator, :*], [:comparison, [:sheet_reference, :"COM.DMD", [:cell, :B30]], [:comparator, :"="], [:sheet_reference, :"COM.DMD", [:cell, :R9]]]], [:arithmetic, [:arithmetic, [:sheet_reference, :"COM.DMD", [:cell, :S28]], [:operator, :*], [:comparison, [:sheet_reference, :"COM.DMD", [:cell, :L6]], [:comparator, :"="], [:sheet_reference, :"COM.DMD", [:cell, :S6]]]], [:operator, :*], [:comparison, [:sheet_reference, :"COM.DMD", [:cell, :B30]], [:comparator, :"="], [:sheet_reference, :"COM.DMD", [:cell, :S9]]]], [:arithmetic, [:arithmetic, [:sheet_reference, :"COM.DMD", [:cell, :T28]], [:operator, :*], [:comparison, [:sheet_reference, :"COM.DMD", [:cell, :L6]], [:comparator, :"="], [:sheet_reference, :"COM.DMD", [:cell, :T6]]]], [:operator, :*], [:comparison, [:sheet_reference, :"COM.DMD", [:cell, :B30]], [:comparator, :"="], [:sheet_reference, :"COM.DMD", [:cell, :T9]]]], [:arithmetic, [:arithmetic, [:sheet_reference, :"COM.DMD", [:cell, :U28]], [:operator, :*], [:comparison, [:sheet_reference, :"COM.DMD", [:cell, :L6]], [:comparator, :"="], [:sheet_reference, :"COM.DMD", [:cell, :U6]]]], [:operator, :*], [:comparison, [:sheet_reference, :"COM.DMD", [:cell, :B30]], [:comparator, :"="], [:sheet_reference, :"COM.DMD", [:cell, :U9]]]], [:arithmetic, [:arithmetic, [:sheet_reference, :"COM.DMD", [:cell, :V28]], [:operator, :*], [:comparison, [:sheet_reference, :"COM.DMD", [:cell, :L6]], [:comparator, :"="], [:sheet_reference, :"COM.DMD", [:cell, :V6]]]], [:operator, :*], [:comparison, [:sheet_reference, :"COM.DMD", [:cell, :B30]], [:comparator, :"="], [:sheet_reference, :"COM.DMD", [:cell, :V9]]]], [:arithmetic, [:arithmetic, [:sheet_reference, :"COM.DMD", [:cell, :W28]], [:operator, :*], [:comparison, [:sheet_reference, :"COM.DMD", [:cell, :L6]], [:comparator, :"="], [:sheet_reference, :"COM.DMD", [:cell, :W6]]]], [:operator, :*], [:comparison, [:sheet_reference, :"COM.DMD", [:cell, :B30]], [:comparator, :"="], [:sheet_reference, :"COM.DMD", [:cell, :W9]]]], [:arithmetic, [:arithmetic, [:sheet_reference, :"COM.DMD", [:cell, :X28]], [:operator, :*], [:comparison, [:sheet_reference, :"COM.DMD", [:cell, :L6]], [:comparator, :"="], [:sheet_reference, :"COM.DMD", [:cell, :X6]]]], [:operator, :*], [:comparison, [:sheet_reference, :"COM.DMD", [:cell, :B30]], [:comparator, :"="], [:sheet_reference, :"COM.DMD", [:cell, :X9]]]], [:arithmetic, [:arithmetic, [:sheet_reference, :"COM.DMD", [:cell, :Y28]], [:operator, :*], [:comparison, [:sheet_reference, :"COM.DMD", [:cell, :L6]], [:comparator, :"="], [:sheet_reference, :"COM.DMD", [:cell, :Y6]]]], [:operator, :*], [:comparison, [:sheet_reference, :"COM.DMD", [:cell, :B30]], [:comparator, :"="], [:sheet_reference, :"COM.DMD", [:cell, :Y9]]]], [:arithmetic, [:arithmetic, [:sheet_reference, :"COM.DMD", [:cell, :Z28]], [:operator, :*], [:comparison, [:sheet_reference, :"COM.DMD", [:cell, :L6]], [:comparator, :"="], [:sheet_reference, :"COM.DMD", [:cell, :Z6]]]], [:operator, :*], [:comparison, [:sheet_reference, :"COM.DMD", [:cell, :B30]], [:comparator, :"="], [:sheet_reference, :"COM.DMD", [:cell, :Z9]]]], [:arithmetic, [:arithmetic, [:sheet_reference, :"COM.DMD", [:cell, :AA28]], [:operator, :*], [:comparison, [:sheet_reference, :"COM.DMD", [:cell, :L6]], [:comparator, :"="], [:sheet_reference, :"COM.DMD", [:cell, :AA6]]]], [:operator, :*], [:comparison, [:sheet_reference, :"COM.DMD", [:cell, :B30]], [:comparator, :"="], [:sheet_reference, :"COM.DMD", [:cell, :AA9]]]], [:arithmetic, [:arithmetic, [:sheet_reference, :"COM.DMD", [:cell, :AB28]], [:operator, :*], [:comparison, [:sheet_reference, :"COM.DMD", [:cell, :L6]], [:comparator, :"="], [:sheet_reference, :"COM.DMD", [:cell, :AB6]]]], [:operator, :*], [:comparison, [:sheet_reference, :"COM.DMD", [:cell, :B30]], [:comparator, :"="], [:sheet_reference, :"COM.DMD", [:cell, :AB9]]]], [:arithmetic, [:arithmetic, [:sheet_reference, :"COM.DMD", [:cell, :AC28]], [:operator, :*], [:comparison, [:sheet_reference, :"COM.DMD", [:cell, :L6]], [:comparator, :"="], [:sheet_reference, :"COM.DMD", [:cell, :AC6]]]], [:operator, :*], [:comparison, [:sheet_reference, :"COM.DMD", [:cell, :B30]], [:comparator, :"="], [:sheet_reference, :"COM.DMD", [:cell, :AC9]]]], [:arithmetic, [:arithmetic, [:sheet_reference, :"COM.DMD", [:cell, :AD28]], [:operator, :*], [:comparison, [:sheet_reference, :"COM.DMD", [:cell, :L6]], [:comparator, :"="], [:sheet_reference, :"COM.DMD", [:cell, :AD6]]]], [:operator, :*], [:comparison, [:sheet_reference, :"COM.DMD", [:cell, :B30]], [:comparator, :"="], [:sheet_reference, :"COM.DMD", [:cell, :AD9]]]], [:arithmetic, [:arithmetic, [:sheet_reference, :"COM.DMD", [:cell, :AE28]], [:operator, :*], [:comparison, [:sheet_reference, :"COM.DMD", [:cell, :L6]], [:comparator, :"="], [:sheet_reference, :"COM.DMD", [:cell, :AE6]]]], [:operator, :*], [:comparison, [:sheet_reference, :"COM.DMD", [:cell, :B30]], [:comparator, :"="], [:sheet_reference, :"COM.DMD", [:cell, :AE9]]]], [:arithmetic, [:arithmetic, [:sheet_reference, :"COM.DMD", [:cell, :AF28]], [:operator, :*], [:comparison, [:sheet_reference, :"COM.DMD", [:cell, :L6]], [:comparator, :"="], [:sheet_reference, :"COM.DMD", [:cell, :AF6]]]], [:operator, :*], [:comparison, [:sheet_reference, :"COM.DMD", [:cell, :B30]], [:comparator, :"="], [:sheet_reference, :"COM.DMD", [:cell, :AF9]]]], [:arithmetic, [:arithmetic, [:sheet_reference, :"COM.DMD", [:cell, :AG28]], [:operator, :*], [:comparison, [:sheet_reference, :"COM.DMD", [:cell, :L6]], [:comparator, :"="], [:sheet_reference, :"COM.DMD", [:cell, :AG6]]]], [:operator, :*], [:comparison, [:sheet_reference, :"COM.DMD", [:cell, :B30]], [:comparator, :"="], [:sheet_reference, :"COM.DMD", [:cell, :AG9]]]], [:arithmetic, [:arithmetic, [:sheet_reference, :"COM.DMD", [:cell, :AH28]], [:operator, :*], [:comparison, [:sheet_reference, :"COM.DMD", [:cell, :L6]], [:comparator, :"="], [:sheet_reference, :"COM.DMD", [:cell, :AH6]]]], [:operator, :*], [:comparison, [:sheet_reference, :"COM.DMD", [:cell, :B30]], [:comparator, :"="], [:sheet_reference, :"COM.DMD", [:cell, :AH9]]]], [:arithmetic, [:arithmetic, [:sheet_reference, :"COM.DMD", [:cell, :AI28]], [:operator, :*], [:comparison, [:sheet_reference, :"COM.DMD", [:cell, :L6]], [:comparator, :"="], [:sheet_reference, :"COM.DMD", [:cell, :AI6]]]], [:operator, :*], [:comparison, [:sheet_reference, :"COM.DMD", [:cell, :B30]], [:comparator, :"="], [:sheet_reference, :"COM.DMD", [:cell, :AI9]]]], [:arithmetic, [:arithmetic, [:sheet_reference, :"COM.DMD", [:cell, :AJ28]], [:operator, :*], [:comparison, [:sheet_reference, :"COM.DMD", [:cell, :L6]], [:comparator, :"="], [:sheet_reference, :"COM.DMD", [:cell, :AJ6]]]], [:operator, :*], [:comparison, [:sheet_reference, :"COM.DMD", [:cell, :B30]], [:comparator, :"="], [:sheet_reference, :"COM.DMD", [:cell, :AJ9]]]], [:arithmetic, [:arithmetic, [:sheet_reference, :"COM.DMD", [:cell, :AK28]], [:operator, :*], [:comparison, [:sheet_reference, :"COM.DMD", [:cell, :L6]], [:comparator, :"="], [:sheet_reference, :"COM.DMD", [:cell, :AK6]]]], [:operator, :*], [:comparison, [:sheet_reference, :"COM.DMD", [:cell, :B30]], [:comparator, :"="], [:sheet_reference, :"COM.DMD", [:cell, :AK9]]]], [:arithmetic, [:arithmetic, [:sheet_reference, :"COM.DMD", [:cell, :AL28]], [:operator, :*], [:comparison, [:sheet_reference, :"COM.DMD", [:cell, :L6]], [:comparator, :"="], [:sheet_reference, :"COM.DMD", [:cell, :AL6]]]], [:operator, :*], [:comparison, [:sheet_reference, :"COM.DMD", [:cell, :B30]], [:comparator, :"="], [:sheet_reference, :"COM.DMD", [:cell, :AL9]]]], [:arithmetic, [:arithmetic, [:sheet_reference, :"COM.DMD", [:cell, :AM28]], [:operator, :*], [:comparison, [:sheet_reference, :"COM.DMD", [:cell, :L6]], [:comparator, :"="], [:sheet_reference, :"COM.DMD", [:cell, :AM6]]]], [:operator, :*], [:comparison, [:sheet_reference, :"COM.DMD", [:cell, :B30]], [:comparator, :"="], [:sheet_reference, :"COM.DMD", [:cell, :AM9]]]], [:arithmetic, [:arithmetic, [:sheet_reference, :"COM.DMD", [:cell, :AN28]], [:operator, :*], [:comparison, [:sheet_reference, :"COM.DMD", [:cell, :L6]], [:comparator, :"="], [:sheet_reference, :"COM.DMD", [:cell, :AN6]]]], [:operator, :*], [:comparison, [:sheet_reference, :"COM.DMD", [:cell, :B30]], [:comparator, :"="], [:sheet_reference, :"COM.DMD", [:cell, :AN9]]]], [:arithmetic, [:arithmetic, [:sheet_reference, :"COM.DMD", [:cell, :AO28]], [:operator, :*], [:comparison, [:sheet_reference, :"COM.DMD", [:cell, :L6]], [:comparator, :"="], [:sheet_reference, :"COM.DMD", [:cell, :AO6]]]], [:operator, :*], [:comparison, [:sheet_reference, :"COM.DMD", [:cell, :B30]], [:comparator, :"="], [:sheet_reference, :"COM.DMD", [:cell, :AO9]]]]]]] 

    r = ReplaceArithmeticOnRangesAst.new
    r.map(input_ast).should == output_ast
  end

end
