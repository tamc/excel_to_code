# frozen_string_literal: true

require_relative 'map_values_to_go'

# The return type when converting
class MapFormulaeToGoResult
  attr_accessor :body
  attr_accessor :body_type

  def initialize
    @body_type = :value
    @counter = 0
    @definitions = []
  end

  def define_variable_with_error(variable_definition)
    result = "v#{@counter += 1}"
    @definitions << "#{result}, err := #{variable_definition}"
    @definitions << 'if err != nil {'
    @definitions << '  return nil, err'
    @definitions << '}'
    result
  end

  def definitions
    return '' if @definitions.empty?

    @definitions.join("\n    ") + "\n    "
  end
end

class MapFormulaeToGo < MapValuesToGo
  attr_accessor :sheet_names
  attr_accessor :getter_method_name
  attr_accessor :result

  FUNCTIONS = {
    :'+' => 'add'
  }.freeze

  def convert(ast)
    @result = MapFormulaeToGoResult.new
    @result.body = map(ast)
    @result
  end

  def number_parameter(ast)
    @result.define_variable_with_error("number(#{map(ast)})")
  end

  def prefix(symbol, ast)
    return map(ast) if symbol == '+'

    function('-', ast)
  end

  def brackets(*contents)
    "(#{contents.map { |a| map(a) }.join(',')})"
  end

  def arithmetic(left, operator, right)
    l = number_parameter(left)
    r = number_parameter(right)
    o = operator.last
    @result.body_type = :function_no_error
    "#{l}#{o}#{r}"
  end

  def comparison(left, operator, right)
    function(operator.last, left, right)
  end

  def function(function_name, *arguments)
    raise NotSupportedException, "Function #{function_name} not supported" unless FUNCTIONS.key?(function_name)

    "#{FUNCTIONS[function_name]}(#{arguments.map { |a| map(a) }.join(',')})"
  end

  def sheet_reference(sheet, reference)
    ref = [sheet, reference.last]
    m = getter_method_name.call(ref)
    "s.#{m}()"
  end
end
