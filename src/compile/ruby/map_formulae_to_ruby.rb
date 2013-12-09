require_relative 'map_values_to_ruby'

class MapFormulaeToRuby < MapValuesToRuby
  
  attr_accessor :sheet_names
  attr_accessor :worksheet
  
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
    'CELL' => 'cell',
    'CHOOSE' => 'choose',
    'CONCATENATE' => 'string_join',
    'COSH' => 'cosh',
    'COUNT' => 'count',
    'COUNTA' => 'counta',
    'FIND' => 'find',
    'HLOOKUP' => 'hlookup',
    'IF' => 'excel_if',
    'IFERROR' => 'iferror',
    'INDEX' => 'index',
    'INT' => 'int',
    'ISNUMBER' => 'isnumber',
    'LARGE' => 'large',
    'LEFT' => 'left',
    'LEN' => 'len',
    'LOG' => 'log',
    'MATCH' => 'excel_match',
    'MAX' => 'max',
    'MID' => 'mid',
    'MIN' => 'min',
    'MMULT' => 'mmult',
    'MOD' => 'mod',
    'PI' => 'pi',
    'PMT' => 'pmt',
    'POWER' => 'power',
    'PV' => 'pv',
    'RANK' => 'rank',
    'RIGHT' => 'right',
    'ROUND' => 'round',
    'ROUNDDOWN' => 'rounddown',
    'ROUNDUP' => 'roundup',
    'SUBSTITUTE' => 'substitute',
    'SUBTOTAL' => 'subtotal',
    'SUM' => 'sum',
    'SUMIF' => 'sumif',
    'SUMIFS' => 'sumifs',
    'SUMPRODUCT' => 'sumproduct',
    'TEXT' => 'text',
    'TRIM' => 'trim',
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
    reference.to_s.downcase.gsub('$','')
  end
  
  def sheet_reference(sheet,reference)
    "#{sheet_names[sheet]}_#{map(reference)}"
  end
  
  def array(*rows)
    "[#{rows.map {|r| map(r)}.join(",")}]"
  end
  
  def row(*cells)
    "[#{cells.map {|r| map(r)}.join(",")}]"
  end
  
end
