require_relative 'extract_formulae'

class ExtractSimpleFormulae < ExtractFormulae

  def start_formula(type,attributes)
    # Simple formulas don't have a type
    return if type
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
    @output[[@sheet_name, @ref]] = ast.to_ast[1]
  end

end
