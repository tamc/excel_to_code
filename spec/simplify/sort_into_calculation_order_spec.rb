require_relative '../spec_helper'

describe SortIntoCalculationOrder do

  it "should be able to turn a set of references into a list sorted into the sequence they need to be calculated" do
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

    calculation_order = [
      [:sheet1,:A3],
      [:sheet1,:A2],
      [:sheet1,:A1],
      [:sheet3,:A1],
      [:sheet2,:A2],
      [:sheet2,:A1],
      [:sheet2,:A3],
      [:sheet3,:A5]
    ]

    sorter = SortIntoCalculationOrder.new
    sorter.sort(references).should == calculation_order
  end # / do


end # / describe
