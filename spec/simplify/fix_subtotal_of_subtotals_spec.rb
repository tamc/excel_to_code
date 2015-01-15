require_relative '../spec_helper'

describe FixSubtotalOfSubtotals do

it "should be able to check if a bit of ast contains a SUBTOTAL function" do

  f = FixSubtotalOfSubtotals.new
  f.is_subtotal?([:number, 1]).should == false
  f.is_subtotal?([:array, [:row, [:function, :SUBTOTAL, [:number, 109], [:array, [:row, [:number, 1]]]]]]).should == true
  f.is_subtotal?([:named_reference, :SUBTOTAL]).should == false
end  

it "should be able to check if a reference refers to a SUBTOTAL function" do
  f = FixSubtotalOfSubtotals.new
  f.references = {
    [:sheet1, :A1] => [:number, 1],
    [:sheet2, :A2] => [:array, [:row, [:function, :SUBTOTAL, [:number, 109], [:array, [:row, [:number, 1]]]]]],
    [:sheet3, :A3] => [:named_reference, :SUBTOTAL]
  }
  f.is_or_refers_to_subtotal?([:number, 1]).should == false
  f.is_or_refers_to_subtotal?([:array, [:row, [:function, :SUBTOTAL, [:number, 109], [:array, [:row, [:number, 1]]]]]]).should == true
  f.is_or_refers_to_subtotal?([:named_reference, :SUBTOTAL]).should == false
  f.is_or_refers_to_subtotal?([:sheet_reference, :sheet1, [:cell, :A1]]).should == false
  f.is_or_refers_to_subtotal?([:sheet_reference, :sheet2, [:cell, :A2]]).should == true
  f.is_or_refers_to_subtotal?([:sheet_reference, :sheet3, [:cell, :A3]]).should == false
  f.is_or_refers_to_subtotal?([:sheet_reference, :sheet10, [:cell, :A9]]).should == false
end

it "should remove references to SUBTOTALS from an array" do
  f = FixSubtotalOfSubtotals.new
  f.references = {
    [:sheet1, :A1] => [:number, 1],
    [:sheet2, :A2] => [:array, [:row, [:function, :SUBTOTAL, [:number, 109], [:array, [:row, [:number, 1]]]]]],
    [:sheet3, :A3] => [:named_reference, :SUBTOTAL]
  }
  f.remove_subtotals_from([:function, :SUM, [:number, 109]]).should == [:function, :SUM, [:number, 109]]
  f.remove_subtotals_from([:function, :SUBTOTAL, [:number, 109]]).should == nil
  f.remove_subtotals_from([:sheet_reference, :sheet2, [:cell, :A2]]).should == nil
  f.remove_subtotals_from([:sheet_reference, :sheet2, [:cell, :A2]]).should == nil
  f.remove_subtotals_from([:array,
                           [:row, [:number, 1]],
                           [:row, [:function, :SUBTOTAL, [:number, 109], [:number, 10] ]],
                           [:row, [:sheet_reference, :sheet1, [:cell, :A1]]],
                           [:row, [:sheet_reference, :sheet2, [:cell, :A2]]]
                          ]).should == [:array,
                           [:row, [:number, 1]],
                           [:row, [:sheet_reference, :sheet1, [:cell, :A1]]]
                          ]
end

it "should prevent SUBTOTAL from counting other SUBTOTALS" do
  f = FixSubtotalOfSubtotals.new
  f.references = {
    [:sheet1, :A1] => [:number, 1],
    [:sheet2, :A2] => [:array, [:row, [:function, :SUBTOTAL, [:number, 109], [:array, [:row, [:number, 1]]]]]],
    [:sheet3, :A3] => [:named_reference, :SUBTOTAL]
  }

  f.map(
    [:function, :SUBTOTAL, [:number, 109], 
      [:array, 
       [:row, [:number, 1]],
       [:row, [:number, 2]]
      ]]
  ).should == 
    [:function, :SUBTOTAL, [:number, 109], 
      [:array, 
       [:row, [:number, 1]],
       [:row, [:number, 2]]
      ]]

  f.count_replaced.should == 0

  f.map(
    [:function, :SUM, 
      [:function, :SUBTOTAL, [:number, 109], 
        [:array, 
          [:row, [:number, 1]],
          [:row, [:function, :SUBTOTAL, [:number, 109], [:array, [:row, [:number, 1]]]]],
          [:row, [:number, 2]]
        ]
    ]]).should == 
      [:function, :SUM, 
        [:function, :SUBTOTAL, [:number, 109], 
          [:array, 
            [:row, [:number, 1]],
            [:row, [:number, 2]]
          ]]]

  f.count_replaced.should == 1

  f.map(
    [:function, :SUBTOTAL, [:number, 109], 
      [:array, 
       [:row, [:number, 1]],
       [:row, [:function, :SUBTOTAL, [:number, 109], [:array, [:row, [:number, 1]]]]],
       [:row, [:number, 2]],
       [:row, [:function, :SUM, [:function, :SUBTOTAL, [:number, 109], [:array, [:row, [:number, 1]]]]]]],
      [:array, 
       [:row, [:number, 1]],
       [:row, [:function, :SUBTOTAL, [:number, 109], [:array, [:row, [:number, 1]]]]],
       [:row, [:number, 2]],
       [:row, [:function, :SUM, [:function, :SUBTOTAL, [:number, 109], [:array, [:row, [:number, 1]]]]],
       [:function, :SUBTOTAL, [:number, 109], [:array, [:row, [:number, 1]]]],
      ]]])
  .should == 
      [:function, :SUBTOTAL, [:number, 109], 
       [:array, 
        [:row, [:number, 1]],
        [:row, [:number, 2]]
       ],
       [:array, 
        [:row, [:number, 1]],
        [:row, [:number, 2]]
      ]]

  f.count_replaced.should == 3

end # / it


end # / describe
