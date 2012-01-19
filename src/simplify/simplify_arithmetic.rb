class SimplifyArithmeticAst
    
  def map(ast)
    return ast unless ast.is_a?(Array)
    operator = ast[0]
    if respond_to?(operator)
      send(operator,*ast[1..-1])
    else
      [operator,*ast[1..-1].map {|a| map(a) }]
    end
  end
  
  def arithmetic(*args)
    return args.first if args.size == 1
    return [:arithmetic,*args] if args.size == 3
    [['^'],['*','/'],['+','-']].each do |op|
      i = args.find_index { |a| a[0] == :operator && op.include?(a[1])}
      next unless i
      pre_operation = i == 1 ? [] : args[0..(i-2)]
      operation = args[(i-1)..(i+1)]
      post_operation = (i + 2) > args.size ? [] : args[(i+2)..-1]
      return arithmetic(*pre_operation,[:arithmetic,*operation],*post_operation)
    end
    return [:arithmetic,*args]
  end
    
end
  

class SimplifyArithmetic
    
  def self.replace(*args)
    self.new.replace(*args)
  end
  
  def replace(input,output)
    rewriter = SimplifyArithmeticAst.new
    input.lines do |line|
      # Looks to match lines with references
      if line =~ /:arithmetic/
        ref, ast = line.split("\t")
        output.puts "#{ref}\t#{rewriter.map(eval(ast)).inspect}"
      else
        output.puts line
      end
    end
  end
end
