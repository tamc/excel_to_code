require_relative 'extract_formulae'

class ExtractSimpleFormulae < ExtractFormulae
  
  def start_formula(type,attributes)
    return if type
    @parsing = true    
    output.write @ref
    output.write "\t"
  end

end
