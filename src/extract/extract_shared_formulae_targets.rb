require_relative 'extract_formulae'

class ExtractSharedFormulaeTargets < ExtractFormulae
  
  attr_accessor :shared_formula_identifier

  def start_formula(type,attributes)
    return unless type == 'shared'
    @shared_formula_identifier = attributes.assoc('si').last
    @parsing = true
  end
  
  def write_formula
    output.write @ref
    output.write "\t"
    output.write @shared_formula_identifier
    output.write "\n"
  end
  
end
