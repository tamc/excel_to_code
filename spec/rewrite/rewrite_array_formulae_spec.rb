require_relative '../spec_helper'

describe RewriteArrayFormulae do

  it "should take array formula like 'reference\\array range\\tformula ast\\n' and output normal formulae" do
    input = { [:Sheet, :B6] => ["B6:B8", [:array, 
                                            [:row, [:arithmetic, [:cell, :"B3"], [:operator, :+], [:cell, :"C3"]]], 
                                            [:row, [:arithmetic, [:cell, :"B4"], [:operator, :+], [:cell, :"C4"]]], 
                                            [:row, [:arithmetic, [:cell, :"B5"], [:operator, :+], [:cell, :"C5"]]]
    ]] }
    RewriteArrayFormulae.rewrite(input).should == {
      [:Sheet, :B6] => [:arithmetic, [:cell, :"B3"], [:operator, :+], [:cell, :"C3"]],
      [:Sheet, :B7] => [:arithmetic, [:cell, :"B4"], [:operator, :+], [:cell, :"C4"]],
      [:Sheet, :B8] => [:arithmetic, [:cell, :"B5"], [:operator, :+], [:cell, :"C5"]]
    }
  end

  it "should work with single cell array formulae" do
    input = { [:Sheet, :B6] => ["B6", [:array, 
                                         [:row, [:arithmetic, [:cell, :"B3"], [:operator, :+], [:cell, :"C3"]]], 
                                         [:row, [:arithmetic, [:cell, :"B4"], [:operator, :+], [:cell, :"C4"]]], 
                                         [:row, [:arithmetic, [:cell, :"B5"], [:operator, :+], [:cell, :"C5"]]]
    ]] }
    RewriteArrayFormulae.rewrite(input).should == {
      [:Sheet, :B6] => [:arithmetic, [:cell, :"B3"], [:operator, :+], [:cell, :"C3"]],
    }
  end

  it "if the array is a single column, should repeat the column across columns in the output" do
    input = { [:Sheet, :B6] => ["B6:C9", [:array, 
                                            [:row, [:arithmetic, [:cell, :"B3"], [:operator, :+], [:cell, :"C3"]]], 
                                            [:row, [:arithmetic, [:cell, :"B4"], [:operator, :+], [:cell, :"C4"]]], 
                                            [:row, [:arithmetic, [:cell, :"B5"], [:operator, :+], [:cell, :"C5"]]]
    ]] }
    RewriteArrayFormulae.rewrite(input).should == {
      [:Sheet, :B6] =>	[:arithmetic, [:cell, :"B3"], [:operator, :+], [:cell, :"C3"]],
      [:Sheet, :B7] =>	[:arithmetic, [:cell, :"B4"], [:operator, :+], [:cell, :"C4"]],
      [:Sheet, :B8] =>	[:arithmetic, [:cell, :"B5"], [:operator, :+], [:cell, :"C5"]],
      [:Sheet, :B9] =>	[:error, :"#N/A"],
      [:Sheet, :C6] =>	[:arithmetic, [:cell, :"B3"], [:operator, :+], [:cell, :"C3"]],
      [:Sheet, :C7] =>	[:arithmetic, [:cell, :"B4"], [:operator, :+], [:cell, :"C4"]],
      [:Sheet, :C8] =>	[:arithmetic, [:cell, :"B5"], [:operator, :+], [:cell, :"C5"]],
      [:Sheet, :C9] =>	[:error, :"#N/A"]
    }
  end

  it "if the array is a single row, should repeat the row across rows in the output" do
    input = { [:Sheet, "B6"] => ["B6:D9", [:array, [:row, [:arithmetic, [:cell, :"B3"], [:operator, :+], [:cell, :"C3"]], [:arithmetic, [:cell, :"B4"], [:operator, :+], [:cell, :"C4"]]] ]] }
    RewriteArrayFormulae.rewrite(input).should == {
      [:Sheet, :B6] =>	[:arithmetic, [:cell, :"B3"], [:operator, :+], [:cell, :"C3"]],
      [:Sheet, :B7] =>	[:arithmetic, [:cell, :"B3"], [:operator, :+], [:cell, :"C3"]],
      [:Sheet, :B8] =>	[:arithmetic, [:cell, :"B3"], [:operator, :+], [:cell, :"C3"]],
      [:Sheet, :B9] =>	[:arithmetic, [:cell, :"B3"], [:operator, :+], [:cell, :"C3"]],
      [:Sheet, :C6] =>	[:arithmetic, [:cell, :"B4"], [:operator, :+], [:cell, :"C4"]],
      [:Sheet, :C7] =>	[:arithmetic, [:cell, :"B4"], [:operator, :+], [:cell, :"C4"]],
      [:Sheet, :C8] =>	[:arithmetic, [:cell, :"B4"], [:operator, :+], [:cell, :"C4"]],
      [:Sheet, :C9] =>	[:arithmetic, [:cell, :"B4"], [:operator, :+], [:cell, :"C4"]],
      [:Sheet, :D6] =>	[:error, :"#N/A"],
      [:Sheet, :D7] =>	[:error, :"#N/A"],
      [:Sheet, :D8] =>	[:error, :"#N/A"],
      [:Sheet, :D9] =>	[:error, :"#N/A"]
    }
  end

  it "if the array is a single cell, should repeat the row across rows and columns in the output" do
    input = { [:Sheet, "B6"] => ["B6:D9", [:array, [:row, [:arithmetic, [:cell, :"B3"], [:operator, :+], [:cell, :"C3"]]] ]] }
    RewriteArrayFormulae.rewrite(input).should == {
      [:Sheet, :B6] => [:arithmetic, [:cell, :"B3"], [:operator, :+], [:cell, :"C3"]],
      [:Sheet, :B7] => [:arithmetic, [:cell, :"B3"], [:operator, :+], [:cell, :"C3"]],
      [:Sheet, :B8] => [:arithmetic, [:cell, :"B3"], [:operator, :+], [:cell, :"C3"]],
      [:Sheet, :B9] => [:arithmetic, [:cell, :"B3"], [:operator, :+], [:cell, :"C3"]],
      [:Sheet, :C6] => [:arithmetic, [:cell, :"B3"], [:operator, :+], [:cell, :"C3"]],
      [:Sheet, :C7] => [:arithmetic, [:cell, :"B3"], [:operator, :+], [:cell, :"C3"]],
      [:Sheet, :C8] => [:arithmetic, [:cell, :"B3"], [:operator, :+], [:cell, :"C3"]],
      [:Sheet, :C9] => [:arithmetic, [:cell, :"B3"], [:operator, :+], [:cell, :"C3"]],
      [:Sheet, :D6] => [:arithmetic, [:cell, :"B3"], [:operator, :+], [:cell, :"C3"]],
      [:Sheet, :D7] => [:arithmetic, [:cell, :"B3"], [:operator, :+], [:cell, :"C3"]],
      [:Sheet, :D8] => [:arithmetic, [:cell, :"B3"], [:operator, :+], [:cell, :"C3"]],
      [:Sheet, :D9] => [:arithmetic, [:cell, :"B3"], [:operator, :+], [:cell, :"C3"]]
    }
  end

  it "should deal with repetition of array formula that only produce a single answer" do
    input = { [:Sheet, :B6] => ["B6:B8", [:function, :SUM, [:array, 
                                                               [:row, [:arithmetic, [:cell, :"B3"], [:operator, :+], [:cell, :"C3"]]], 
                                                               [:row, [:arithmetic, [:cell, :"B4"], [:operator, :+], [:cell, :"C4"]]], 
                                                               [:row, [:arithmetic, [:cell, :"B5"], [:operator, :+], [:cell, :"C5"]]]
    ]]] }
    RewriteArrayFormulae.rewrite(input).should == {
      [:Sheet, :B6] => [:function, :SUM, [:array, [:row, [:arithmetic, [:cell, :"B3"], [:operator, :+], [:cell, :"C3"]]], [:row, [:arithmetic, [:cell, :"B4"], [:operator, :+], [:cell, :"C4"]]], [:row, [:arithmetic, [:cell, :"B5"], [:operator, :+], [:cell, :"C5"]]]]],
      [:Sheet, :B7] => [:function, :SUM, [:array, [:row, [:arithmetic, [:cell, :"B3"], [:operator, :+], [:cell, :"C3"]]], [:row, [:arithmetic, [:cell, :"B4"], [:operator, :+], [:cell, :"C4"]]], [:row, [:arithmetic, [:cell, :"B5"], [:operator, :+], [:cell, :"C5"]]]]],
      [:Sheet, :B8] => [:function, :SUM, [:array, [:row, [:arithmetic, [:cell, :"B3"], [:operator, :+], [:cell, :"C3"]]], [:row, [:arithmetic, [:cell, :"B4"], [:operator, :+], [:cell, :"C4"]]], [:row, [:arithmetic, [:cell, :"B5"], [:operator, :+], [:cell, :"C5"]]]]]
    }
  end

  it "should deal with functions that may potentially return arrays" do
    input = { [:Sheet, :B6] => ["B6:B8", [:function, :INDEX, [:array, [:row, [:number, 1]], [:row, [:number, 2]], [[:row, [:number, 3]]]], [:null], [:number, 1]]] }
    RewriteArrayFormulae.rewrite(input).should == {
      [:Sheet, :B6] => [:function, :INDEX, [:function, :INDEX, [:array, [:row, [:number, 1]], [:row, [:number, 2]], [[:row, [:number, 3]]]], [:null], [:number, 1]], [:number, 1], [:number, 1]],
      [:Sheet, :B7] => [:function, :INDEX, [:function, :INDEX, [:array, [:row, [:number, 1]], [:row, [:number, 2]], [[:row, [:number, 3]]]], [:null], [:number, 1]], [:number, 2], [:number, 1]],
      [:Sheet, :B8] => [:function, :INDEX, [:function, :INDEX, [:array, [:row, [:number, 1]], [:row, [:number, 2]], [[:row, [:number, 3]]]], [:null], [:number, 1]], [:number, 3], [:number, 1]]
    }
  end

end
