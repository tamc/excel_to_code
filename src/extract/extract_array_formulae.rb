require_relative 'extract_formulae'

class ExtractArrayFormulae < ExtractFormulae
  
  attr_accessor :array_range
  
  def start_formula(type,attributes)
    return unless type == 'array' && attributes.assoc('ref')
    @array_range = attributes.assoc('ref').last
    @parsing = true
  end
  
  def write_formula
    return false if @formula.empty?
    output.write @ref
    output.write "\t"
    output.write @array_range
    output.write "\t"
    output.write @formula.join.gsub(/[\n\r]+/,'')
    output.write "\n"
  end

end
