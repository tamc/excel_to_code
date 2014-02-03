class SimplifyArithmeticAst
  
  attr_accessor :map_count  

  def map(ast)
    @map_count = 0
    @brackets_to_remove = []
    simplify_arithmetic(ast)
    remove_brackets(ast)
  end

  def simplify_arithmetic(ast)
    @map_count += 1
    return ast unless ast.is_a?(Array)
    ast.each { |a| simplify_arithmetic(a) }
    case ast[0]
    when :arithmetic; arithmetic(ast)
    when :brackets; @brackets_to_remove << ast
    end
    ast
  end
  
  def remove_brackets(ast)
    @brackets_to_remove.uniq.each do |ast|
      raise NotSupportedException.new("Multiple arguments not supported in brackets #{ast.inspect}") if ast.size > 2
      ast.replace(ast[1])
    end
    ast
  end
  
  def arithmetic(ast)
    return unless ast.size > 4
    # This sets the operator precedence
    i = nil
    [{:'^' => 1},{:'*' => 2,:'/' => 2},{:'+' => 3,:'-' => 3}].each do |op|
      i = ast.find_index { |a| a[0] == :operator && op.has_key?(a[1])}
      break if i
    end
    if i
      # Now we need to wrap that operation in its own arithmetic clause
      old_clause = ast[(i-1)..(i+1)]
      # Make sure we do any mapping
      #old_clause.each { |a| map(a) }
      # Now create a new clause
      new_clause = [:arithmetic, *old_clause]
      # And insert it back into the ast
      ast[(i-1)..(i+1)]  = [new_clause]
      # Redo the mapping
      simplify_arithmetic(ast)
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
