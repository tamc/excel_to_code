require_relative 'extract_formulae'

class ExtractSharedFormulae < ExtractFormulae
  
  def start_formula(type,attributes)
    return unless type == 'shared' && attributes.assoc('ref')
    @shared_range = attributes.assoc('ref').last
    @shared_formula_identifier = attributes.assoc('si').last
    @parsing = true
  end
  
  def write_formula
    return if @formula.empty?
    formula_text = @formula.join.gsub(/[\r\n]+/,'')
    ast = Formula.parse(formula_text)
    unless ast
      $stderr.puts "Could not parse #{@sheet_name} #{@ref} #{formula_text}"
      exit
    end
    # FIXME: Should leave in original form rather than converting to ast?
    @output[[@sheet_name, @ref]] = [@shared_range, @shared_formula_identifier, ast.to_ast[1]]
  end
  
end
