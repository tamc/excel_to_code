require_relative 'extract_formulae'

class ExtractSharedFormulaeTargets < ExtractFormulae
  
  def start_formula(type,attributes)
    return unless type == 'shared'
    @shared_formula_identifier = attributes.assoc('si').last
    @parsing = true
  end
  
  def write_formula
    @output[[@sheet_name, @ref]] = @shared_formula_identifier
  end
  
end
