class SimplifyArithmeticAst
  
  def map(ast)
    @brackets_to_remove = []
    simplify_arithmetic(ast)
    remove_brackets
    ast
  end

  def simplify_arithmetic(ast)
    return ast unless ast.is_a?(Array)
    ast.each { |a| simplify_arithmetic(a) }
    case ast[0]
    when :arithmetic; arithmetic(ast)
    when :brackets; @brackets_to_remove << ast
    end
  end
  
  def remove_brackets
    @brackets_to_remove.uniq.each do |ast|
      raise NotSupportedException.new("Multiple arguments not supported in brackets #{ast.inspect}") if ast.size > 2
      ast.replace(ast[1])
    end
  end

  # This sets the operator precedence
  OPERATOR_PRECEDENCE = [
    {:'^' => 1},
    {:'*' => 2,:'/' => 2},
    {:'+' => 3,:'-' => 3}
  ]
  
  def arithmetic(ast)
    # If smaller than 4, will only be a simple operation (e.g., 1+1 or 2*4)
    # If more than 4, will be like 1+2*3 and so needs turning into 1+(2*3)
    return unless ast.size > 4
    OPERATOR_PRECEDENCE.each do |op|
      i = nil
      while i = ast.find_index { |a| a[0] == :operator && op.has_key?(a[1])}
        # Now we need to wrap that operation in its own arithmetic clause
        old_clause = ast[(i-1)..(i+1)]
        # Now create a new clause
        new_clause = [:arithmetic, *old_clause]
        # And insert it back into the ast
        ast[(i-1)..(i+1)]  = [new_clause]
        # Redo the mapping
      end
    end
    # FIXME: this feels like a bodge
    if ast.size == 2 && ast[0] == :arithmetic && ast[1].is_a?(Array) && ast[1][0] == :arithmetic
      ast.replace(ast[1])
    end
  end
end
  

class SimplifyArithmetic
    
  def self.replace(*args)
    self.new.replace(*args)
  end
  
  def replace(input,output)
    rewriter = SimplifyArithmeticAst.new
    input.each_line do |line|
      # Looks to match lines with references
      if line =~ /:arithmetic/
        content = line.split("\t")
        ast = eval(content.pop)
        output.puts "#{content.join("\t")}\t#{rewriter.map(ast).inspect}"
      else
        output.puts line
      end
    end
  end
end
