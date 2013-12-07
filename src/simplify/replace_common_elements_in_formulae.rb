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

  VALUES = {:number => true, :string => true, :blank => true, :null => true, :error => true, :boolean_true => true, :boolean_false => true, :sheet_reference => true, :cell => true, :row => true}
  
  def replace_repeated_formulae(ast)
    return ast unless ast.is_a?(Array)
    return ast if VALUES.has_key?(ast.first)    
    replacement = @common_elements[ast]
    if replacement
      ast.replace(replacement)
    else
      ast.each { |a| replace_repeated_formulae(a) }
    end
  end   

end
