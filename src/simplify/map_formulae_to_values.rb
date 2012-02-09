require_relative '../compile'
require_relative '../excel/excel_functions'
require_relative '../util'

class FormulaeCalculator
  include ExcelFunctions
end

class MapFormulaeToValues
  
  def initialize
    @value_for_ast = MapValuesToRuby.new
    @calculator = FormulaeCalculator.new
  end
  
  def map(ast)
    return ast unless ast.is_a?(Array)
    operator = ast[0]
    if respond_to?(operator)
      send(operator,*ast[1..-1])
    else
      [operator,*ast[1..-1].map {|a| map(a) }]
    end
  end
  
  def prefix(operator,argument)
    argument_value = value(map(argument))
    return [:prefix, operator, map(argument)] if argument_value == :not_a_value
    return ast_for_value(argument_value || 0) if operator == "+"
    ast_for_value((argument_value || 0) * -1)
  end
  
  def arithmetic(left,operator,right)
    l = value(map(left))
    r = value(map(right))
    if (l != :not_a_value) && (r != :not_a_value)
      formula_value(operator.last,l,r)
    else
      [:arithmetic,map(left),operator,map(right)]
    end
  end
  
  alias :comparison :arithmetic
  
  def percentage(number)
    ast_for_value(value([:percentage, number]))
  end
  
  def string_join(*args)
    values = args.map { |a| value(map(a)) } # FIXME: These eval statements are really bugging me. Must find a better solution
    if values.any? { |a| a == :not_a_value }
      [:string_join,*args.map { |a| map(a) }]
    else
      ast_for_value(@calculator.string_join(*values))
    end
  end
  
  FUNCTIONS_THAT_SHOULD_NOT_BE_CONVERTED = %w{TODAY RAND RANDBETWEEN INDIRECT}
  
  def function(name,*args)
    if FUNCTIONS_THAT_SHOULD_NOT_BE_CONVERTED.include?(name)
      [:function,name,*args.map { |a| map(a) }]
    elsif respond_to?("map_#{name.downcase}")
      send("map_#{name.downcase}",*args)
    else
      values = args.map { |a| value(map(a)) }
      if values.any? { |a| a == :not_a_value }
        [:function,name,*args.map { |a| map(a) }]
      else
        formula_value(name,*values)
      end
    end
  end
  
  def map_index(array,row_number,column_number = nil)
    array_mapped = map(array)
    row_as_number = value(map(row_number))
    column_as_number = (value(map(column_number)) || nil) if column_number
    if column_number
      if row_as_number == :not_a_value || column_as_number == :not_a_value
        return [:function, "INDEX", array_mapped, map(row_number), map(column_number)]
      end
    else
      if row_as_number == :not_a_value
        return [:function, "INDEX", array_mapped, map(row_number)]
      end
    end
    case array_mapped.first
    when :array
      array_as_values = array_mapped[1..-1].map do |row|
        row[1..-1].map do |cell|
          cell
        end
      end 
    when :cell, :sheet_reference, :blank, :number, :percentage, :string, :error, :boolean_true, :boolean_false
      array_as_values = [[array_mapped]]
    else
      if column_number
        return  [:function, "INDEX", array_mapped, map(row_number), map(column_number)]
      else
        return  [:function, "INDEX", array_mapped, map(row_number)]
      end
    end
    result = @calculator.send(MapFormulaeToRuby::FUNCTIONS["INDEX"],array_as_values,row_as_number,column_as_number)
    result = ast_for_value(result) unless result.is_a?(Array)
    result
  end
    
  def value(ast)
    return extract_values_from_array(ast) if ast.first == :array
    return :not_a_value unless @value_for_ast.respond_to?(ast.first)
    eval(@value_for_ast.send(*ast))
  end
  
  def extract_values_from_array(ast)
    ast[1..-1].map do |row|
      row[1..-1].map do |cell|
        v = value(cell)
        return :not_a_value if v == :not_a_value
        v
      end
    end 
  end
  
  def formula_value(ast_name,*arguments)
    raise NotSupportedException.new("#{ast_name.inspect} function not recognised in #{MapFormulaeToRuby::FUNCTIONS.inspect}") unless MapFormulaeToRuby::FUNCTIONS.has_key?(ast_name)
    ast_for_value(@calculator.send(MapFormulaeToRuby::FUNCTIONS[ast_name],*arguments))
  end
  
  def ast_for_value(value)
    case value
    when Numeric; [:number,value.inspect]
    when true; [:boolean_true]
    when false; [:boolean_false]
    when Symbol; 
      raise NotSupportedException.new("Error #{value.inspect} not recognised") unless MapFormulaeToRuby::REVERSE_ERRORS[value.inspect]
      [:error,MapFormulaeToRuby::REVERSE_ERRORS[value.inspect]]
    when String; [:string,value]
    when Array; [:array,*value.map { |row| [:row, *row.map { |c| ast_for_value(c) }]}]
    when nil; [:blank]
    else
      raise NotSupportedException.new("Ast for #{value.inspect} of class #{value.class} not recognised")
    end
  end
  
end
