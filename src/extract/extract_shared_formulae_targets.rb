require_relative 'extract_formulae'

class ExtractSharedFormulaeTargets < ExtractFormulae
  
  def start_formula(type,attributes)
    return unless type == 'shared'
    @parsing = true
  end
  
  def write_formula
    output.write @ref
    output.write "\n"
  end
  
end
