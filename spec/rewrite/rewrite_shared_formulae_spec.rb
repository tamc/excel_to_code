require_relative '../spec_helper'

describe RewriteSharedFormulae do

  it "should take shared formula hash, a shared_target hash, and ouput a set of formula hashes" do
    formula_shared = { 
      ["Sheet", "B3"] => ["B3:C4", "0", [:cell,'A1']],
      ["Sheet", "Z5"] => ["Z5","1", [:cell,'A1']]
    }
    formula_shared_targets = {
      ["Sheet", "B3"] => "0",
      ["Sheet", "B4"] => "0",
      ["Sheet", "C3"] => "0",
      ["Sheet", "C4"] => "0",
      ["Sheet", "Z5"] => "1"
    }
    RewriteSharedFormulae.rewrite(formula_shared, formula_shared_targets).should == {
      ["Sheet", "B3"]  =>	[:cell, "A1"],
      ["Sheet", "B4"]  =>	[:cell, "A2"],
      ["Sheet", "C3"]  =>	[:cell, "B1"],
      ["Sheet", "C4"]  =>	[:cell, "B2"],
      ["Sheet", "Z5"]  =>	[:cell, "A1"]
    }
  end

  it "should take cope with occasions when the shared formula is not the top left in the range" do
    formula_shared = { 
      ["Sheet", "C3"] => ["B3:C4", "0", [:cell,'B1']],
    }
    formula_shared_targets = {
      ["Sheet", "B3"] => "0",
      ["Sheet", "B4"] => "0",
      ["Sheet", "C3"] => "0",
      ["Sheet", "C4"] => "0",
      ["Sheet", "Z5"] => "0"
    }
    RewriteSharedFormulae.rewrite(formula_shared, formula_shared_targets).should == {
      ["Sheet", "B3"]  =>	[:cell, "A1"],
      ["Sheet", "B4"]  =>	[:cell, "A2"],
      ["Sheet", "C3"]  =>	[:cell, "B1"],
      ["Sheet", "C4"]  =>	[:cell, "B2"]
    }
  end

  it "should cope with exceptions in the shared formula range" do
    formula_shared = { 
      ["Sheet", "B3"] => ["B3:C4", "0", [:cell,'A1']],
      ["Sheet", "Z5"] => ["Z5","1", [:cell,'A1']]
    }
    formula_shared_targets = {
      ["Sheet", "B3"] => "0",
      ["Sheet", "C3"] => "0",
      ["Sheet", "C4"] => "0",
      ["Sheet", "Z5"] => "1"
    }
    RewriteSharedFormulae.rewrite(formula_shared, formula_shared_targets).should == {
      ["Sheet", "B3"]  =>	[:cell, "A1"],
      ["Sheet", "C3"]  =>	[:cell, "B1"],
      ["Sheet", "C4"]  =>	[:cell, "B2"],
      ["Sheet", "Z5"]  =>	[:cell, "A1"]
    }
  end

  it "should cope with overlapping shared formula ranges" do
    formula_shared = { 
      ["Sheet", "B3"] => ["B3:C4", "0", [:cell,'A1']],
      ["Sheet", "B4"] => ["B3:C4", "1", [:cell,'C2']]
    }
    formula_shared_targets = {
      ["Sheet", "B3"] => "0",
      ["Sheet", "B4"] => "1",
      ["Sheet", "C3"] => "1",
      ["Sheet", "C4"] => "0",
    }
    RewriteSharedFormulae.rewrite(formula_shared, formula_shared_targets).should == {
      ["Sheet", "B3"]  =>	[:cell, "A1"],
      ["Sheet", "B4"]  =>	[:cell, "C2"],
      ["Sheet", "C3"]  =>	[:cell, "D1"],
      ["Sheet", "C4"]  =>	[:cell, "B2"]
    }
  end
end
