class SimplifyArithmeticAst
    
  def map(ast)
    return ast unless ast.is_a?(Array)
    case ast[0]
    when :brackets; brackets(ast)
    when :arithmetic; arithmetic(ast)
    end
    ast.each { |a| map(a) }
    ast
  end
  
  def brackets(ast)
    raise NotSupportedException.new("Multiple arguments not supported in brackets #{args.inspect}") if ast.size > 2
    ast.replace(map(ast[1]))
  end
  
  def arithmetic(ast)
    case ast.size
    when 2; return map(ast[1]) # Not really arithmetic
    when 4; # Normal arithmetic that doesn't need re-arranging
      ast.each { |a| map(a) }
    else
      # This sets the operator precedence
      i = nil
      [{'^' => 1},{'*' => 2,'/' => 2},{'+' => 3,'-' => 3}].each do |op|
        i = ast.find_index { |a| a[0] == :operator && op.has_key?(a[1])}
        break if i
      end
      if i
        # Now we need to wrap that operation in its own arithmetic clause
        old_clause = ast[(i-1)..(i+1)]
        # Make sure we do any mapping
        old_clause.each { |a| map(a) }
        # Now create a new clause
        new_clause = [:arithmetic, *old_clause]
        # And insert it back into the ast
        ast[(i-1)..(i+1)]  = [new_clause]
        # Redo the mapping
        map(ast)
      end
    end
    ast.each { |a| map(a) }
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
