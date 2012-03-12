require_relative 'map_values_to_c'

class MapFormulaeToC < MapValuesToC
  
  attr_accessor :sheet_names
  attr_accessor :worksheet
  attr_reader :initializers
  attr_reader :counter
  
  def initialize
    reset
  end
  
  def reset
    @initializers = []
    @counter = 0
  end
  
  FUNCTIONS = {
    '*' => 'multiply',
    '+' => 'add',
    '-' => 'subtract',
    '/' => 'divide',
    '<' => 'less_than?',
    '<=' => 'less_than_or_equal?',
    '<>' => 'not_equal?',
    '=' => 'excel_equal?',
    '>' => 'more_than?',
    '>=' => 'more_than_or_equal?',
    'ABS' => 'abs',
    'AND' => 'excel_and',
    'AVERAGE' => 'average',
    'CHOOSE' => 'choose',
    'COSH' => 'cosh',
    'COUNT' => 'count',
    'COUNTA' => 'counta',
    'FIND' => 'find',
    'IF' => 'excel_if',
    'IFERROR' => 'iferror',
    'INDEX' => 'index',
    'LEFT' => 'left',
    'MATCH' => 'excel_match',
    'MAX' => 'max',
    'MIN' => 'min',
    'MOD' => 'mod',
    'PI' => 'pi',
    'PMT' => 'pmt',
    'ROUND' => 'round',
    'ROUNDDOWN' => 'rounddown',
    'ROUNDUP' => 'roundup',
    'SUBTOTAL' => 'subtotal',
    'SUM' => 'sum',
    'SUMIF' => 'sumif',
    'SUMIFS' => 'sumifs',
    'SUMPRODUCT' => 'sumproduct',
    'VLOOKUP' => 'vlookup',
    '^' => 'power'
  }
  
  def prefix(symbol,ast)
    return map(ast) if symbol == "+"
    return "negative(#{map(ast)})"
  end
  
  def brackets(*contents)
    "(#{contents.map { |a| map(a) }.join(',')})"
  end
  
  def arithmetic(left,operator,right)
    "#{FUNCTIONS[operator.last]}(#{map(left)},#{map(right)})"
  end
  
  def string_join(*strings)
    "string_join(#{strings.map {|a| map(a)}.join(',')})"
  end
  
  def comparison(left,operator,right)
    "#{FUNCTIONS[operator.last]}(#{map(left)},#{map(right)})"
  end
  
  def function(function_name,*arguments)
    if FUNCTIONS.has_key?(function_name)
      "#{FUNCTIONS[function_name]}(#{arguments.map { |a| map(a) }.join(",")})"
    else
      raise NotSupportedException.new("Function #{function_name} not supported")
    end
  end
  
  def cell(reference)
    reference.downcase.gsub('$','')
  end
  
  def sheet_reference(sheet,reference)
    return map(reference) if worksheet && worksheet == sheet
    "#{sheet_names[sheet]}_#{map(reference).downcase}()"
  end

  def array(*rows)
    # Make sure we get the right dimensions
    number_of_rows = rows.size
    number_of_columns = rows.max { |r| r.size }.size - 1
    
    # First we have to create an excel array
    array_name = "array#{@counter}"

    cells = rows.map do |r|
      r.shift if r.first == :row
      r.map do |c| 
        map(c)
      end
    end.flatten.join(',')
    
    initializers << "ExcelValue #{array_name}[] = {#{cells}};"
    
    # Then we need to assign it to an excel value
    range_name = "range#{@counter}"
    initializers << "ExcelValue #{range_name} = new_excel_range(#{array_name},#{number_of_rows},#{number_of_columns});"

    @counter +=1

    range_name
  end


end
