require_relative '../spec_helper'

describe ReplaceCommonElementsInFormulae do

  it "should work through formulae, replacing elements that are common to other formulae with a cell reference" do

    input = {
      ["sheet1", "A1"] => [:function, "INDEX", [:array, [:row, [:cell, "A1"], [:cell, "A2"], [:cell, "A3"]]], [:number, 2]]
    }

    common = {
      [:array, [:row, [:cell, "A1"], [:cell, "A2"], [:cell, "A3"]]] => [:cell, "common0"]
    }

    expected_output = {
      ["sheet1", "A1"] => [:function, "INDEX", [:cell, "common0"], [:number, 2]]
    }

    r = ReplaceCommonElementsInFormulae.new
    r.replace(input,common).should == expected_output
  end # /it

end # /Describe
