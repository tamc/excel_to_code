require_relative '../spec_helper'

describe CachingFormulaParser do
  
  it "should turn sheet and cell references into symbols" do
    CachingFormulaParser.parse("Sheet1!B$1").should == [:sheet_reference, :Sheet1, [:cell, :"B$1"]]
    CachingFormulaParser.parse("Sheet1!B$1:B10").should == [:sheet_reference, :Sheet1, [:area, :"B$1", :B10]]
  end

  it "should turn numbers into numeric values" do
    CachingFormulaParser.parse("1.31").should == [:number, 1.31]
    CachingFormulaParser.parse("10%").should == [:number, 0.1]
  end

  it "should turn operators and comparators into symbols" do
    CachingFormulaParser.parse("1+1").should == [:arithmetic, [:number, 1], [:operator, :+], [:number, 1]]
    CachingFormulaParser.parse("1>=1").should == [:comparison, [:number, 1], [:comparator, :'>='], [:number, 1]]
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

  it "should cache sheet references" do 
    first = CachingFormulaParser.parse("Sheet1!B$1")
    second = CachingFormulaParser.parse("Sheet1!B$1")
    first.object_id.should == second.object_id
  end

  it "should turn function names into symbols" do 
    CachingFormulaParser.parse('INDIRECT("A1")').should == [:function, :INDIRECT, [:string, "A1"]]

  end

  it "should turn sheet references that refer to an error (e.g., Control!#REF! into the error)" do
    CachingFormulaParser.parse('Control!#REF!').should == [:error, :'#REF!']

  end

  it "should remove percentages straight away" do
    CachingFormulaParser.parse("15.15%").should == [:number, 0.1515]
    CachingFormulaParser.parse("ROUND(24.35)%").should == [:arithmetic, [:function, :ROUND, [:number, 24.35]], [:operator, :"/"], [:number, 100.0]]

  end 

  it "should catch external references" do
    expect { CachingFormulaParser.parse("[1]Sheet1!B1") }.to raise_error(ExternalReferenceException)
    expect { CachingFormulaParser.parse("'[1] Output tab'!B1") }.to raise_error(ExternalReferenceException)
    CachingFormulaParser.parse("'a [1] Output tab'!B1").should == [:sheet_reference, :"a [1] Output tab", [:cell, :B1]]



  end

  it "should report failures to parse" do
    expect { CachingFormulaParser.parse("{NOT PARSABLE}") }.to raise_error(ParseFailedException)
  end

end

