require_relative 'extract_formulae'

class ExtractSharedFormulae < ExtractFormulae
  
  attr_accessor :shared_range
  
  def start_formula(type,attributes)
    return unless type == 'shared' && attributes.assoc('ref')
    @shared_range = attributes.assoc('ref').last
    @parsing = true
  end
  
  def write_formula
    return if @formula.empty?
    output.write @ref
    output.write "\t"
    output.write @shared_range
    output.write "\t"
    output.write @formula.join.gsub(/[\n\r]+/,'')
    output.write "\n"
  end
  
end
