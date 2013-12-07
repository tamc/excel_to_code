require_relative '../excel'

class WrapFormulaeThatReturnArraysAndAReNotInArraysAst
  def map(ast)
    return ast unless ast.is_a?(Array)
    function(ast) if ast.first == :function
    ast
  end

  # Only does MMULT at the moment
  FORMULAE_THAT_RETURN_ARRAYS = { "MMULT" => true }
  
  def function(ast)
    return unless FORMULAE_THAT_RETURN_ARRAYS.has_key?(ast[1])
    ast.replace( [:function, "INDEX", ast.dup,  [:number, "1"], [:number, "1"]])
  end
end

class WrapFormulaeThatReturnArraysAndAReNotInArrays
  
  def self.replace(*args)
    self.new.replace(*args)
  end
  
  def replace(input,output)
    r = WrapFormulaeThatReturnArraysAndAReNotInArraysAst.new

    input.each_line do |line|
      # Looks to match lines that contain formulae that return ranges, such as MMULT
      content = line.split("\t")
      ast = eval(content.pop)
      output.puts "#{content.join("\t")}\t#{r.map(ast).inspect}"          
    end
  end
  
end
