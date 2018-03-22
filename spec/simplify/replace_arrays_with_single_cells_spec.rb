require_relative '../spec_helper'

describe ReplaceArraysWithSingleCellsAst do
  
  it "should replace array literals (e.g., {A1,B1;A2,B2}) with the appropriate cell (e.g., A1) where it is required" do

  r = ReplaceArraysWithSingleCellsAst.new

  ast = [:array, [:row, [:sheet_reference, :"sheet1", [:cell, :"B1"]], [:sheet_reference, :"sheet1", [:cell, :"C1"]]]]

  r.ref = [:sheet1, :B10]
  r.map(ast).should == [:sheet_reference, :"sheet1", [:cell, :"B1"]] 

  r.ref = [:sheet1, :C10]
  r.map(ast).should == [:sheet_reference, :"sheet1", [:cell, :"C1"]] 

  r.ref = [:sheet1, :D10]
  r.map(ast).should ==   [:error, :"#VALUE!"]  # [:array, [:row, [:sheet_reference, :sheet1, [:cell, :B1]], [:sheet_reference, :sheet1, [:cell, :C1]]]] # Excel would return a #VALUE! but we need to return the array in case it is part of function that accepts an array FIXME: Do this properly someitme

  r.ref = [:sheet1, :B10]
  r.map([:function, :SUM, *ast]).should == [:function, :SUM, *ast]

  ast_vertical =  [:array, [:row, [:sheet_reference, :"sheet1", [:cell, :"A2"]]], [:row, [:sheet_reference, :"sheet1", [:cell, :"A3"]]]]

  r.ref = [:sheet1, :Z2]
  r.map(ast_vertical).should == [:sheet_reference, :sheet1, [:cell, :A2]]
  
  r.ref = [:sheet1, :Z12]
  r.map(ast_vertical).should == [:error, :"#VALUE!"] 

  r.ref = [:sheet1, :Z2]
  r.map([:function, :SUM, *ast_vertical]).should == [:function, :SUM, *ast_vertical]

  r.ref = [:sheet1, :B2]
  r.map([:arithmetic, ast, [:operator, :+], ast_vertical]).should == [:arithmetic, [:sheet_reference, :"sheet1", [:cell, :"B1"]], [:operator, :+], [:sheet_reference, :sheet1, [:cell, :A2]]]

  r.map([:string_join, ast, ast_vertical]).should == [:string_join, [:sheet_reference, :"sheet1", [:cell, :"B1"]], [:sheet_reference, :sheet1, [:cell, :A2]]]

  r.map([:function, :INDIRECT, [:string_join, ast, ast_vertical]]).should == [:function, :INDIRECT, [:string_join, [:sheet_reference, :"sheet1", [:cell, :"B1"]], [:sheet_reference, :sheet1, [:cell, :A2]]]]

  sumifast = [:function, :SUMIF, ast_vertical, ast_vertical ] 
  sumifast_result = [:function, :SUMIF, ast_vertical, [:sheet_reference, :sheet1, [:cell, :A3]]]
  r.ref = [:sheet1, :B3]
  r.map(sumifast).should == sumifast_result

  sumifsast = [:function, :SUMIFS, ast_vertical, ast_vertical, ast_vertical ] 
  sumifsast_result = [:function, :SUMIFS, ast_vertical, ast_vertical, [:sheet_reference, :sheet1, [:cell, :A3]]]
  r.ref = [:sheet1, :B3]
  r.map(sumifsast).should == sumifsast_result


  ast_vertical =  [:array, [:row, [:sheet_reference, :"sheet1", [:cell, :"A2"]]], [:row, [:sheet_reference, :"sheet1", [:cell, :"A3"]]]]
  if_ast = [:function, :IF, [:boolean_true], ast_vertical, ast_vertical]
  r.ref = [:sheet1, :B3]
  r.map(if_ast).should == [:function, :IF, [:boolean_true], [:sheet_reference, :sheet1, [:cell, :A3]], [:sheet_reference, :sheet1, [:cell, :A3]]]

  index_ast = [:function, :INDEX, ast_vertical, ast_vertical, ast_vertical]
  r.ref = [:sheet1, :B3]
  r.map(index_ast).should == [:function, :INDEX, ast_vertical, [:sheet_reference, :sheet1, [:cell, :A3]], [:sheet_reference, :sheet1, [:cell, :A3]]]

  end

end
