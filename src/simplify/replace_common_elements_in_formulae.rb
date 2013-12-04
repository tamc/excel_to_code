class ReplaceCommonElementsInFormulae
    
  def self.replace(*args)
    self.new.replace(*args)
  end
  
  attr_accessor :common_elements
  
  def replace(input, common_elements)
    @common_elements = common_elements
    input.each do |ref, ast|
      replace_repeated_formulae(ast)
    end
    input
  end

  VALUES = [:number,:string,:blank,:null,:error,:boolean_true,:boolean_false,:sheet_reference,:cell, :row]
  
  def replace_repeated_formulae(ast)
    return ast unless ast.is_a?(Array)
    return ast if VALUES.include?(ast.first)    
    replacement = @common_elements[ast]
    if replacement
      ast.replace(replacement)
    else
      ast.each { |a| replace_repeated_formulae(a) }
    end
  end   

end
