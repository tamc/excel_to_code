require_relative '../spec_helper'

describe CachingFormulaParser do
  
  it "should turn sheet and cell references into symbols" do
    CachingFormulaParser.parse("Sheet1!B$1").should == [:sheet_reference, :Sheet1, [:cell, :"B$1"]]
    CachingFormulaParser.parse("Sheet1!B$1:B10").should == [:sheet_reference, :Sheet1, [:area, :"B$1", :B10]]
  end

  it "should turn numbers into numeric values" do
    CachingFormulaParser.parse("1.31").should == [:number, 1.31]
    CachingFormulaParser.parse("10%").should == [:percentage, 10]
  end

  it "should cache numbers and strings and booleans and blanks and errors and operators and comparators" do
    first = CachingFormulaParser.parse("1.31")
    second = CachingFormulaParser.parse("1.31")
    first.object_id.should == second.object_id

    first = CachingFormulaParser.parse('"Hello"')
    second = CachingFormulaParser.parse('"Hello"')
    first.object_id.should == second.object_id

    first = CachingFormulaParser.parse('TRUE')
    second = CachingFormulaParser.parse('TRUE')
    first.object_id.should == second.object_id

    first = CachingFormulaParser.parse('FALSE')
    second = CachingFormulaParser.parse('FALSE')
    first.object_id.should == second.object_id

    first = CachingFormulaParser.parse('#DIV/0!')
    second = CachingFormulaParser.parse('#DIV/0!')
    first.object_id.should == second.object_id

    first = CachingFormulaParser.parse('1 + 1')
    second = CachingFormulaParser.parse('1 + 1')
    first[1].object_id.should == second[1].object_id
    first[2].object_id.should == second[2].object_id
    first[3].object_id.should == second[3].object_id

    first = CachingFormulaParser.parse('1 < 1')
    second = CachingFormulaParser.parse('1 < 1')
    first[1].object_id.should == second[1].object_id
    first[2].object_id.should == second[2].object_id
    first[3].object_id.should == second[3].object_id
  end

  it "should turn function names into symbols" do 
    CachingFormulaParser.parse('INDIRECT("A1")').should == [:function, :INDIRECT, [:string, "A1"]]

  end

end

