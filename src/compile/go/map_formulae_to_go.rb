# frozen_string_literal: true

require_relative 'map_values_to_go'

class MapFormulaeToGo < MapValuesToGo
  attr_accessor :sheet_names
  attr_accessor :getter_method_name
  attr_accessor :definitions
  attr_accessor :result_type

  FUNCTIONS = {
    :'+' => 'add'
  }.freeze

  def convert(ast)
    @result_type = :value
    @counter = 0
    @definitions = []
    map(ast)
  end

  def get_definitions
    return "" if @definitions.empty?
    @definitions.join("\n    ") + "\n    "
  end

  def number_parameter(ast)
    value = map(ast)
    variable = "v#{@counter += 1}"
    @definitions << "#{variable}, err := number(#{value})"
    @definitions << 'if err != nil {'
    @definitions << '  return nil, err'
    @definitions << '}'
    variable
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
    @result_type = :function_no_error
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
