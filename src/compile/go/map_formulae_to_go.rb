# frozen_string_literal: true

require_relative 'map_values_to_go'

class MapFormulaeToGo < MapValuesToGo
  attr_accessor :sheet_names
  attr_accessor :getter_method_name

  FUNCTIONS = {
    :'+' => 'add'
  }.freeze

  def prefix(symbol, ast)
    return map(ast) if symbol == '+'

    function('-', ast)
  end

  def brackets(*contents)
    "(#{contents.map { |a| map(a) }.join(',')})"
  end

  def arithmetic(left, operator, right)
    function(operator.last, left, right)
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
