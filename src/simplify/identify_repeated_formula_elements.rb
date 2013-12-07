class IdentifyRepeatedFormulaElements
  
  attr_accessor :references
  attr_accessor :counted_elements
  attr_accessor :bothered_threshold
  
  def initialize
    @references = {}
    @counted_elements = {}
    @counted_elements.default_proc = lambda do |hash,key|
      hash[key] = 0
    end
    @bothered_threshold = 20
  end
  
  def count(references)
    @references = references
    references.each do |ref,ast|
      identify_repeated_formulae(ast)
    end
    return @counted_elements
  end

  IGNORE_TYPES = {:number => true, :string => true, :blank => true, :null => true, :error => true, :boolean_true => true, :boolean_false => true, :sheet_reference => true, :cell => true,  :row => true}
  
  def identify_repeated_formulae(ast)
    string = ast.to_s
    return unless ast.is_a?(Array)
    return if IGNORE_TYPES.has_key?(ast.first)
    return if string.length < bothered_threshold
    @counted_elements[ast] += 1
    ast.each do |a|
      identify_repeated_formulae(a)
    end
  end   
end
