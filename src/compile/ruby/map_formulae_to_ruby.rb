require_relative 'map_values_to_ruby'

class MapFormulaeToRuby < MapValuesToRuby
  
  FUNCTIONS = {
    '+' => 'add',
    '-' => 'subtract',
    '*' => 'multiply',
    '/' => 'divide',
    '^' => 'power',
    'COSH' => 'cosh',
    'PI' => 'pi',
    'SUM' => 'sum'
  }
  
  def arithmetic(left,operator,right)
    "#{FUNCTIONS[operator.last]}(#{map(left)},#{map(right)})"
  end
  
  def function(function_name,*arguments)
    if FUNCTIONS.has_key?(function_name)
      "#{FUNCTIONS[function_name]}(#{arguments.map { |a| map(a) }.join(",")})"
    else
      raise NotSupportedException.new("Function #{function_name} not supported")
    end
  end
  
end
