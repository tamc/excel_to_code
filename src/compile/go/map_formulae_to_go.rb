# frozen_string_literal: true

require_relative 'map_values_to_go'

# The return type when converting
class MapFormulaeToGoResult
  attr_accessor :body
  attr_accessor :body_type
  attr_accessor :return_type

  def initialize
    @body_type = :value
    @return_type = :interface
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

  def error!
    @return_type = :error
    @body_type = :error_value
  end

  def number!
    @return_type = :number
    @body_type = :value
  end

  def boolean!
    @return_type = :boolean
    @body_type = :value
  end

  def string!
    @return_type = :string
    @body_type = :value
  end

  def blank!
    @return_type = :blank
    @body_type = :value
  end

  def interface!
    @return_type = :interface
    @body_type = :value
  end

  def function_no_error!(returns_type:)
    @return_type = returns_type
    @body_type = :function_no_error
  end

  def number?
    @return_type == :number
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
    v = map(ast)
    return v if @result.number?

    @result.define_variable_with_error("number(#{v})")
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
    @result.function_no_error!(returns_type: :number)
    "(#{l}#{o}#{r})"
  end

  def comparison(left, operator, right)
    function(operator.last, left, right)
  end

  def function(function_name, *arguments)
    raise NotSupportedException, "Function #{function_name} not supported" unless FUNCTIONS.key?(function_name)

    "#{FUNCTIONS[function_name]}(#{arguments.map { |a| map(a) }.join(',')})"
  end

  def sheet_reference(sheet, reference)
    @result.interface!
    ref = [sheet, reference.last]
    m = getter_method_name.call(ref)
    "s.#{m}()"
  end
end
