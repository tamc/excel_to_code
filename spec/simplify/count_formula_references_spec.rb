require_relative '../spec_helper'

describe CountFormulaReferences do

  it "should be able to count the number of times a formula is referenced" do
    references = {
      [:sheet1, :A1] => [:cell, :"$A$2"],
      [:sheet1, :A2] => [:cell, :A3],
      [:sheet1, :A3] => [:number, 1],
      [:sheet2, :A1] => [:cell, :A2],
      [:sheet2, :A2] => [:sheet_reference,:sheet3,[:cell,:A1]],
      [:sheet2, :A3] => [:cell, :A2],
      [:sheet3, :A1] => [:number, 5],
      [:sheet3, :A5] => [:number, 10]    
    }

    reference_count = {
      [:sheet1, :A1] => 0,
      [:sheet1, :A2] => 1,
      [:sheet1, :A3] => 1,
      [:sheet2, :A1] => 0,
      [:sheet2, :A2] => 2,
      [:sheet2, :A3] => 0,
      [:sheet3, :A1] => 1,
      [:sheet3, :A5] => 0
    }

    counter = CountFormulaReferences.new
    counter.count(references).should == reference_count
  end # / do


end # / describe
