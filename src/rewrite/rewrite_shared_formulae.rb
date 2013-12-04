require_relative 'ast_copy_formula'

class RewriteSharedFormulae
  def self.rewrite(*args)
    new.rewrite(*args)
  end
  
  def rewrite(formula_shared, formula_shared_targets)
    @output = {}
    @formula_shared_targets = formula_shared_targets

    formula_shared.each do |ref, a|
      copy_range = a[0]
      shared_formula_identifier = a[1]
      shared_ast = a[2]
      share_formula(ref, shared_ast, copy_range, shared_formula_identifier)
    end
    @output
  end
  
  def share_formula(ref, shared_ast, copy_range, shared_formula_identifier)
    copier = AstCopyFormula.new
    copy_range = Area.for(copy_range)
    copy_range.calculate_excel_variables
    
    start_reference = copy_range.excel_start
    
    r = Reference.for(ref.last)
    r.calculate_excel_variables
    
    offset_from_formula_to_start_rows = start_reference.excel_row_number - r.excel_row_number
    offset_from_formula_to_start_columns = start_reference.excel_column_number - r.excel_column_number
    
    copy_range.offsets.each do |row,column|
      new_ref = [ref.first, start_reference.offset(row,column)]
      next unless @formula_shared_targets.include?(new_ref)
      next unless shared_formula_identifier == @formula_shared_targets[new_ref]
      copier.rows_to_move = row + offset_from_formula_to_start_rows
      copier.columns_to_move = column + offset_from_formula_to_start_columns
      ast = copier.copy(shared_ast)
      @output[new_ref] = ast
    end
  end
  
end
