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
  
  def brackets(*args)
    raise NotSupportedException.new("Multiple arguments not supported in brackets #{args.inspect}") if args.size > 1
    map(args.first)
  end
  
  def arithmetic(*args)
    return map(args.first) if args.size == 1
    return [:arithmetic,*args.map {|a| map(a) }] if args.size == 3
    [['^'],['*','/'],['+','-']].each do |op|
      i = args.find_index { |a| a[0] == :operator && op.include?(a[1])}
      next unless i
      pre_operation = i == 1 ? [] : args[0..(i-2)]
      operation = args[(i-1)..(i+1)]
      post_operation = (i + 2) > args.size ? [] : args[(i+2)..-1]
      return arithmetic(*pre_operation.map {|a| map(a) },[:arithmetic,*operation],*post_operation.map {|a| map(a) })
    end
    return [:arithmetic,*args.map {|a| map(a) }]
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
