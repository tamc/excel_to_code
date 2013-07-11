require_relative 'ast_copy_formula'

class RewriteSharedFormulae
  def self.rewrite(input, shared_targets, output)
    new.rewrite(input, shared_targets, output)
  end
  
  def rewrite(input, shared_targets, output)
    shared_targets = Hash[*shared_targets.each_line.map { |l| l.split("\t").map(&:strip) }.flatten]
    input.each_line do |line|
      ref, copy_range,  shared_formula_identifier, formula = line.split("\t")
      share_formula(ref, formula, copy_range, shared_formula_identifier, shared_targets, output)
    end
  end
  
  def share_formula(ref, formula, copy_range, shared_formula_identifier, shared_targets, output)
    shared_ast = eval(formula)
    copier = AstCopyFormula.new
    copy_range = Area.for(copy_range)
    copy_range.calculate_excel_variables
    
    start_reference = copy_range.excel_start
    
    r = Reference.for(ref)
    r.calculate_excel_variables
    
    offset_from_formula_to_start_rows = start_reference.excel_row_number - r.excel_row_number
    offset_from_formula_to_start_columns = start_reference.excel_column_number - r.excel_column_number
    
    copy_range.offsets.each do |row,column|
      new_ref = start_reference.offset(row,column)
      next unless shared_targets.include?(new_ref)
      next unless shared_formula_identifier == shared_targets[new_ref]
      copier.rows_to_move = row + offset_from_formula_to_start_rows
      copier.columns_to_move = column + offset_from_formula_to_start_columns
      ast = copier.copy(shared_ast)
      output.puts "#{new_ref}\t#{ast.inspect}"
    end
  end
  
end
