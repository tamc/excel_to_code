require_relative 'map_values_to_ruby'

class MapFormulaeToRuby < MapValuesToRuby
  
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
  
end
