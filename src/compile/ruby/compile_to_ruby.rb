class MapAstToRuby

  def map(ast)
    if ast.is_a?(Array)
      operator = ast.shift
      if respond_to?(operator)
        send(operator,*ast)
      else
        [operator,*ast.map {|a| map(a) }].join('')
      end
    else
      return ast
    end
  end
  
  FUNCTIONS = {
    '+' => 'add',
    '-' => 'subtract',
    '*' => 'multiply',
    '/' => 'divide',
    '^' => 'power',
  }
  
  def arithmetic(left,operator,right)
    "#{FUNCTIONS[operator.last]}(#{map(left)},#{map(right)})"
  end
  
  def number(text)
    text
  end
  
end

class CompileToRuby
  
  def self.rewrite(*args)
    self.new.rewrite(*args)
  end
  
  def rewrite(input,output)
    mapper = MapAstToRuby.new
    input.lines do |line|
      ref, formula = line.split("\t")
      output.puts "  def #{ref.downcase}; #{mapper.map(eval(formula))}; end"
    end
  end
  
end