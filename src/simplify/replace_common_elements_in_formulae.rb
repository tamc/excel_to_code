class ReplaceCommonElementsInFormulae
    
  def self.replace(*args)
    self.new.replace(*args)
  end
  
  attr_accessor :common_elements
  
  def replace(input,common,output)
    @common_elements ||= {}
    common.readlines.map do |a|
      ref, element = a.split("\t")
      @common_elements[element.strip] = [:cell, ref]
    end
    input.lines do |line|
      ref, formula = line.split("\t")
      output.puts "#{ref}\t#{replace_repeated_formulae(eval(formula)).inspect}"
    end
  end
  
  def replace_repeated_formulae(ast)
    return ast unless ast.is_a?(Array)
    return ast if [:number,:string,:blank,:null,:error,:boolean_true,:boolean_false,:sheet_reference,:cell, :row].include?(ast.first)    
    string = ast.inspect
    return ast if string.length < 20
    if @common_elements.has_key?(string)
      return @common_elements[string]
    end
    ast.map do |a|
      replace_repeated_formulae(a)
    end
  end   

end
