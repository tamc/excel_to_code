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
  
  def arithmetic(left,operator,right)
    l = value(map(left))
    r = value(map(right))
    if (l != :not_a_value) && (r != :not_a_value)
      formula_value(operator.last,l,r)
    else
      [:arithmetic,left,operator,right]
    end
  end
  
  alias :comparison :arithmetic
  
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
    else
      values = args.map { |a| value(map(a)) }
      if values.any? { |a| a == :not_a_value }
        [:function,name,*args.map { |a| map(a) }]
      else
        formula_value(name,*values)
      end
    end
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
    when Symbol; [:error,MapFormulaeToRuby::REVERSE_ERRORS[value.inspect]]
    when String; [:string,value]
    when Array; [:array,*value.map { |row| [:row, *row.map { |c| ast_for_value(c) }]}]
    when nil; [:blank]
    else
      raise NotSupportedException.new("Ast for #{value.inspect} of class #{value.class} not recognised")
    end
  end
  
end
