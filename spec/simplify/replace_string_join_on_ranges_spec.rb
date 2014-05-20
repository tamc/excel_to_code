require_relative '../spec_helper'

describe ReplaceStringJoinOnRangesAST do

  it "should replace string joins on ranges (e.g., A1:A2&B1:B2 = [A1&B1, A2&B2]). Note Excel only seems to do this if formulae are entered in array form?" do

    r = ReplaceStringJoinOnRangesAST.new
    r.map([:string_join, [:string, "A"], [:number, 1]]).should == [:string_join, [:string, "A"], [:number, 1]]
                                                                  r.map([:string_join, [:string, "A"], [:array, [:row, [:number, 1]], [:row, [:number, 2]]]]).should == [:array, [:row, [:string_join, [:string, "A"], [:number, 1]]], [:row, [:string_join, [:string, "A"], [:number, 2]]]]
                                                                  r.map([:string_join, [:array, [:row, [:string, "A"]], [:row, [:string, "B"]]], [:array, [:row, [:number, 1]], [:row, [:number, 2]]]]).should == [:array, [:row, [:string_join, [:string, "A"], [:number, 1]]], [:row, [:string_join, [:string, "B"], [:number, 2]]]]
                                                                  r.map([:string_join, [:array, [:row, [:string, "A"]], [:row, [:string, "B"]]], [:number, 1]]).should ==[:array, [:row, [:string_join, [:string, "A"], [:number, 1]]], [:row, [:string_join, [:string, "B"], [:number, 1]]]]
                                                                  r.map([:string_join, [:array, [:row, [:string, "A"]], [:row, [:string, "B"]]], [:number, 1], [:array, [:row, [:number, 1]], [:row, [:number, 2]]]]).should == [:array, [:row, [:string_join, [:string, "A"], [:number, 1], [:number, 1]]], [:row, [:string_join, [:string, "B"], [:number, 1], [:number, 2]]]]
  end

end
