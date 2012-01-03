require_relative 'ast_copy_formula'

class RewriteSharedFormulae
  def self.rewrite(input,output)
    new.rewrite(input,output)
  end
  
  def rewrite(input,output)
    input.lines do |line|
      ref, copy_range, formula = line.split("\t")
      share_formula(formula,copy_range,output)
    end
  end
  
  def share_formula(formula,copy_range,output)
    shared_ast = eval(formula)
    copier = AstCopyFormula.new
    copy_range = Area.for(copy_range)
    copy_range.calculate_excel_variables
    start_reference = copy_range.excel_start
    
    copy_range.offsets.each do |row,column|
      ref = start_reference.offset(row,column)
      copier.rows_to_move = row
      copier.columns_to_move = column
      ast = copier.copy(shared_ast)
      output.puts "#{ref}\t#{ast.inspect}"
    end
  end
  
end