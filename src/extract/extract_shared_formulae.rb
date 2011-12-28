require_relative 'extract_formulae'

class ExtractSharedFormulae < ExtractFormulae
  
  def start_formula(type,attributes)
    return unless type == 'shared' && attributes.assoc('ref')
    @parsing = true    
    output.write @ref
    output.write "\t"
    output.write attributes.assoc('ref').last
    output.write "\t"
  end
  
end
