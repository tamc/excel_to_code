require_relative '../spec_helper'

describe ReplaceReferencesToBlanksWithZeros do

  it "If a cell contains a reference, and only a reference, to a cell that is blank or doesn't exist, replace it with a zero" do

    references = {
      [:sheet1, :A1] => [:sheet_reference, :sheet1, [:cell, :A2]],
      [:sheet1, :B1] => [:sheet_reference, :sheet1, [:cell, :B2]],
      [:sheet1, :B2] => [:blank],
      [:sheet1, :C1] => [:function, :SUM, [:sheet_reference, :sheet1, [:cell, :A2]]],
    }

    r = ReplaceReferencesToBlanksWithZeros.new(references, :sheet1)

    references.each { |ref, ast| r.map(ast) }

    references.should == {
      [:sheet1, :A1] => [:number, 0],
      [:sheet1, :B1] => [:number, 0],
      [:sheet1, :B2] => [:blank],
      [:sheet1, :C1] => [:function, :SUM, [:sheet_reference, :sheet1, [:cell, :A2]]],
    }

  end

  it "If should be replacing with a zero, but can't because can't inline, then wrap in a check" do

    blankref = [:sheet_reference, :sheet1, [:cell, :A2]] 

    references = {
      [:sheet1, :A1] => blankref.dup,
    }

    inline_ast = lambda { |sheet, ref, references| false }

    r = ReplaceReferencesToBlanksWithZeros.new(references, :sheet1, inline_ast)

    references.each { |ref, ast| r.map(ast) }

    references.should == {
      [:sheet1, :A1] => [:function, :IF, 
                         [:function, :ISBLANK, blankref],
                         [:number, 0],
                         blankref],
    }

  end

end
