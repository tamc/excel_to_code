class RewriteArrayFormulae
  def self.rewrite(input,output)
    new.rewrite(input,output)
  end
  
  def rewrite(input,output)
    input.lines do |line|
      ref, array_range, formula = line.split("\t")
      array_formula(formula,array_range,output)
    end
  end
  
  def array_formula(formula,array_range,output)
    array_ast = eval(formula)
    array_range = Area.for(array_range)
    array_range.calculate_excel_variables
    start_reference = array_range.excel_start
    
    # First we rewrite the master array formula
    array_ast = [:formula,[:function,"ARRAYFORMULA",*array_ast[1..-1]]]
    output.puts "#{start_reference}\t#{array_ast.inspect}"
    
    # Then we rewrite each of the subsidiaries
    array_range.offsets.each do |row,column|
      next if row == 0 && column == 0
      ref = start_reference.offset(row,column)
      ast = [:formula, [:function,"CONTINUE", [:cell, start_reference], row+1, column+1]]
      output.puts "#{ref}\t#{ast.inspect}"
    end
  end
  
end