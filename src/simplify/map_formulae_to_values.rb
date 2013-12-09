require_relative '../compile'
require_relative '../excel/excel_functions'
require_relative '../util'

class FormulaeCalculator
  include ExcelFunctions
  attr_accessor :original_excel_filename
end

class MapFormulaeToValues
  
  attr_accessor :original_excel_filename
  attr_accessor :replacements_made_in_the_last_pass
  
  def initialize
    @value_for_ast = MapValuesToRuby.new
    @calculator = FormulaeCalculator.new
    @replacements_made_in_the_last_pass = 0
  end

  def original_excel_filename=(new_filename)
    @original_excel_filename = new_filename
    @calculator.original_excel_filename = new_filename
  end

  def reset
    # Not used any more
    # FIXME: Remove references to this method
  end

  DO_NOT_MAP = {:number => true, :string => true, :blank => true, :null => true, :error => true, :boolean_true => true, :boolean_false => true, :sheet_reference => true, :cell => true}

  def map(ast)
    ast[1..-1].each do |a| 
      next unless a.is_a?(Array)
      next if DO_NOT_MAP[(a[0])]
      map(a)
    end # Depth first best in this case?
    send(ast[0], ast) if respond_to?(ast[0])
    ast
  end

  # [:prefix, operator, argument]
  def prefix(ast)
    operator, argument = ast[1], ast[2]
    argument_value = value(argument)
    return if argument_value == :not_a_value
    return ast.replace(ast_for_value(argument_value || 0)) if operator == "+"
    ast.replace(ast_for_value(@calculator.negative(argument_value)))
  end
  
  # [:arithmetic, left, operator, right]
  def arithmetic(ast)
    left, operator, right = ast[1], ast[2], ast[3]
    l = value(left)
    r = value(right)
    return if (l == :not_a_value) || (r == :not_a_value)
    ast.replace(formula_value(operator.last,l,r))
  end
  
  alias :comparison :arithmetic

  # [:percentage, number]
  def percentage(ast)
    ast.replace(ast_for_value(value([:percentage, ast[1]])))
  end
  
  # [:string_join, stringA, stringB, ...]
  def string_join(ast)
    values = ast[1..-1].map { |a| value(a) } 
    return if values.any? { |a| a == :not_a_value }
    ast.replace(ast_for_value(@calculator.string_join(*values)))
  end
  
  # [:function, function_name, arg1, arg2, ...]
  def function(ast)
    name = ast[1]
    return if name == :INDIRECT
    return if name == :OFFSET
    return if name == :COLUMN
    if respond_to?("map_#{name.to_s.downcase}")
      send("map_#{name.to_s.downcase}",ast)
    else
      values = ast[2..-1].map { |a| value(a) }
      return if values.any? { |a| a == :not_a_value }
      ast.replace(formula_value(name,*values))
    end
  end

  # [:function, "COUNT", range]
  def map_count(ast)
    range = ast[2]
    return unless [:array, :cell, :sheet_reference].include?(range.first)
    range = array_as_values(range)
    ast.replace(ast_for_value(range.size * range.first.size))
  end
  
  # [:function, "INDEX", array, row_number, column_number]
  def map_index(ast)
    return map_index_with_only_two_arguments(ast) if ast.length == 4

    array_mapped = ast[2] 
    row_as_number = value(ast[3])
    column_as_number = value(ast[4])

    return if row_as_number == :not_a_value 
    return if column_as_number == :not_a_value

    array_as_values = array_as_values(array_mapped)
    return unless array_as_values

    result = @calculator.send(MapFormulaeToRuby::FUNCTIONS[:INDEX],array_as_values,row_as_number,column_as_number)
    result = [:number, 0] if result == [:blank]
    result = ast_for_value(result)
    ast.replace(result)
  end
  
  # [:function, "INDEX", array, row_number]
  def map_index_with_only_two_arguments(ast)
    array_mapped = ast[2]
    row_as_number = value(ast[3])
    return if row_as_number == :not_a_value
    array_as_values = array_as_values(array_mapped)
    return unless array_as_values
    result = @calculator.send(MapFormulaeToRuby::FUNCTIONS[:INDEX],array_as_values,row_as_number)
    result = [:number, 0] if result == [:blank]
    result = ast_for_value(result)
    ast.replace(result)
  end
  
  def array_as_values(array_mapped)
    case array_mapped.first
    when :array
      array_mapped[1..-1].map do |row|
        row[1..-1].map do |cell|
          cell
        end
      end 
    when :cell, :sheet_reference, :blank, :number, :percentage, :string, :error, :boolean_true, :boolean_false
      [[array_mapped]]
    else
      nil
    end
  end

  ERRORS = {
    :"#NAME?" => :name,
    :"#VALUE!" => :value,
    :"#DIV/0!" => :div0,
    :"#REF!" => :ref,
    :"#N/A" => :na,
    :"#NUM!" => :num
  }
    
  def value(ast)
    return extract_values_from_array(ast) if ast.first == :array
    case ast.first
    when :blank; nil
    when :null; nil
    when :number; ast[1]
    when :percentage; ast[1]/100.0
    when :string; ast[1]
    when :error; ERRORS[ast[1]]
    when :boolean_true; true
    when :boolean_false; false
    else return :not_a_value
    end
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
    raise NotSupportedException.new("#{ast_name} function not recognised in #{MapFormulaeToRuby::FUNCTIONS.inspect}") unless MapFormulaeToRuby::FUNCTIONS.has_key?(ast_name)
    ast_for_value(@calculator.send(MapFormulaeToRuby::FUNCTIONS[ast_name],*arguments))
  end
  
  def ast_for_value(value)
    return value if value.is_a?(Array) && value.first.is_a?(Symbol)
    @replacements_made_in_the_last_pass += 1
    ast = case value
    when Numeric; [:number,value]
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
    CachingFormulaParser.map(ast)
  end
  
end
