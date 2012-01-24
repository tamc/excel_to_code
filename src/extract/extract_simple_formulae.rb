require_relative 'extract_formulae'

class ExtractSimpleFormulae < ExtractFormulae
  
  def start_formula(type,attributes)
    return if type
    @parsing = true
  end
  
  def write_formula   
    return if @formula.empty?
    output.write @ref
    output.write "\t"
    output.write @formula.join
    output.write "\n"
  end

end
