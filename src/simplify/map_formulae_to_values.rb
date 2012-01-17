require_relative '../compile'
require_relative '../excel/excel_functions'

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
    l = value(left)
    r = value(right)
    if l && r
      formula_value(operator.last,l,r)
    else
      [:arithmetic,left,operator,right]
    end
  end
  
  alias :comparison :arithmetic
  
  FUNCTIONS_THAT_CHANGE_AT_RUNTIME = %w{TODAY RAND RANDBETWEEN}
  
  def function(name,*args)
    if FUNCTIONS_THAT_CHANGE_AT_RUNTIME.include?(name)
      [:function,name,*args]
    else
      values = args.map { |a| value(a) }
      if values.any? { |a| a.nil? }
        [:function,name,*args]
      else
        formula_value(name,*values)
      end
    end
  end
  
  def value(ast)
    return nil unless @value_for_ast.respond_to?(ast.first)
    @value_for_ast.send(*ast)
  end
  
  def formula_value(ast_name,*arguments)
    ast_for_value(@calculator.send(MapFormulaeToRuby::FUNCTIONS[ast_name],*arguments))
  end
  
  def ast_for_value(value)
    case value
    when Numeric; [:number,value.to_s]
    when true; [:boolean_true]
    when false; [:boolean_false]
    when Symbol; [:error,value.to_s]
    when String; [:string,value]
    else value
    end
  end
  
end
