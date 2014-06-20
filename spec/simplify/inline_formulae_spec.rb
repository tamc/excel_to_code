require_relative '../spec_helper'

describe InlineFormulaeAst do

  it "should replace references to other cell with the contents of that cell" do

    references = {
      [:sheet1, :A2] => [:cell, :"A3"]
    }
    r = InlineFormulaeAst.new(references, :sheet1)
    r.map([:cell, :A2]).should == [:cell, :A3]
    r.map([:sheet_reference, :sheet1, [:cell, :A2]]).should == [:cell, :A3]
    r.map([:function, :sum, [:sheet_reference, :sheet1, [:cell, :A2]]]).should == [:function, :sum, [:cell, :A3]]
  end

  it "should not replace references to otehr cells when they are used as arguments in OFFSET, ROW and COLUMN functions" do
    references = {
      [:sheet1, :A2] => [:cell, :"A3"]
    }
    r = InlineFormulaeAst.new(references, :sheet1)
    r.map([:function, :ROW, [:cell, :A2]]).should == [:function, :ROW, [:cell, :A2]]
    r.map([:function, :COLUMN, [:cell, :A2]]).should == [:function, :COLUMN, [:cell, :A2]]
    r.map([:function, :OFFSET, [:cell, :A2], [:cell, :A2]]).should == [:function, :OFFSET, [:cell, :A2], [:cell, :A3]]
  end


  it "should accept a block, which can be used to decide whether to inline a particular reference or not" do

    references = {
      [:sheet1, :A2] => [:string, "Yes"],
      [:sheet1, :B1] => [:string, "No"]
    }
    inline_ast_decision = lambda do |sheet, cell, references|
      cell != :B1
    end
    r = InlineFormulaeAst.new(references, :sheet1, inline_ast_decision)
    r.map([:string_join, [:cell, :A2], [:cell, :B1]]).should == [:string_join, [:string, "Yes"], [:cell, :B1]]
  end

  it "If the reference refers to a cell that doesn't exist in the Excel input, add the missing cell to the references hash as a [:blank] cell" do

    ast = [:cell, :A2]

    references = {
      [:sheet1, :A1] => ast
    }

    current_sheet_name = :sheet1

    r = InlineFormulaeAst.new(references, current_sheet_name)

    r.map(ast).should == [:inlined_blank]

    references.should == {
      [:sheet1, :A1] => [:inlined_blank],
      [:sheet1, :A2] => [:blank]
    }

  end
end

