require_relative '../spec_helper'

describe IdentifyRepeatedFormulaElements do

  it "should be able to count the number of times a formula is referenced" do
    formulae = {
      ["sheet1", 'A1'] => [:function, "INDEX", [:array, [:row, [:cell, "A1"], [:cell, "A2"], [:cell, "A3"]]], [:number, 2]],
      ["sheet1", 'A2'] => [:function, "INDEX", [:array, [:row, [:cell, "A1"], [:cell, "A2"], [:cell, "A3"]]], [:number, 20]],
      ["sheet1", 'A3'] => [:function, "INDEX", [:array, [:row, [:cell, "A1"], [:cell, "A2"], [:cell, "A3"]]], [:number, 1], [:number, 2]],
      ["sheet2", 'A1'] => [:function, "INDEX", [:array, [:row, [:cell, "A1"], [:cell, "A2"], [:cell, "A3"]]], [:number, 2]],
      ["sheet2", 'A3'] => [:cell, "A2"],
      ["sheet3", 'A1'] => [:number, 5],
      ["sheet3", 'A5'] => [:number, 10]    
    }

    count = {
      [:function, "INDEX", [:array, [:row, [:cell, "A1"], [:cell, "A2"], [:cell, "A3"]]], [:number, 2]] => 2,
      [:array, [:row, [:cell, "A1"], [:cell, "A2"], [:cell, "A3"]]] => 4,
      [:function, "INDEX", [:array, [:row, [:cell, "A1"], [:cell, "A2"], [:cell, "A3"]]], [:number, 20]] => 1,
      [:function, "INDEX", [:array, [:row, [:cell, "A1"], [:cell, "A2"], [:cell, "A3"]]], [:number, 1], [:number, 2]] => 1,
    }


    identifier = IdentifyRepeatedFormulaElements.new
    identifier.count(formulae).should == count
  end # / do


end # / describe
