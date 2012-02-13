class IdentifyRepeatedFormulaElements
  
  attr_accessor :references
  attr_accessor :counted_elements
  attr_accessor :current_sheet
  attr_accessor :bothered_threshold
  
  def initialize
    @references = {}
    @counted_elements = {}
    @counted_elements.default_proc = lambda do |hash,key|
      h = {}
      h.default_proc = lambda { |hash,key| hash[key] = 0 }
      hash[key] = h
    end
    @bothered_threshold = 20
    @current_sheet = ""
  end
  
  def count(references)
    @references = references
    references.each do |sheet,cells|
      @current_sheet = sheet
      cells.each do |ref,ast|
        identify_repeated_formulae(ast)
      end
    end
    return @counted_elements
  end

  
  def identify_repeated_formulae(ast)
    string = ast.to_s
    return unless ast.is_a?(Array)
    return if [:number,:string,:blank,:null,:error,:boolean_true,:boolean_false,:sheet_reference,:cell, :row].include?(ast.first)
    return if string.length < bothered_threshold
    @counted_elements[@current_sheet][string] += 1
    ast.each do |a|
      identify_repeated_formulae(a)
    end
  end   
end
