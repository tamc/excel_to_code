require_relative 'map_values_to_ruby'

class MapFormulaeToRuby < MapValuesToRuby
  
  attr_accessor :sheet_names
  attr_accessor :worksheet
  
  FUNCTIONS = {
    :'*' => 'multiply',
    :'+' => 'add',
    :'-' => 'subtract',
    :'/' => 'divide',
    :'<' => 'less_than?',
    :'<=' => 'less_than_or_equal?',
    :'<>' => 'not_equal?',
    :'=' => 'excel_equal?',
    :'>' => 'more_than?',
    :'>=' => 'more_than_or_equal?',
    :'ABS' => 'abs',
    :'ADDRESS' => 'address',
    :'AND' => 'excel_and',
    :'AVERAGE' => 'average',
    :'AVERAGEIFS' => 'averageifs',
    :'CELL' => 'cell',
    :'CHAR' => 'char',
    :'CHOOSE' => 'choose',
    :'CONCATENATE' => 'string_join',
    :'COSH' => 'cosh',
    :'COUNT' => 'count',
    :'COUNTA' => 'counta',
    :'COUNTIF' => 'countif',
    :'ENSURE_IS_NUMBER' => 'ensure_is_number',
    :'EXP' => 'exp',
    :'FIND' => 'find',
    :'FORECAST' => 'forecast',
    :'HLOOKUP' => 'hlookup',
    :'HYPERLINK' => 'hyperlink',
    :'IF' => 'excel_if',
    :'IFERROR' => 'iferror',
    :'INDEX' => 'index',
    :'INT' => 'int',
    :'ISBLANK' => 'isblank',
    :'ISERR' => 'iserr',
    :'ISERROR' => 'iserror',
    :'ISNUMBER' => 'isnumber',
    :'LARGE' => 'large',
    :'LEFT' => 'left',
    :'LEN' => 'len',
    :'LN' => 'ln',
    :'LOG' => 'log',
    :'LOG10' => 'log',
    :'LOWER' => 'lower',
    :'MATCH' => 'excel_match',
    :'MAX' => 'max',
    :'MID' => 'mid',
    :'MIN' => 'min',
    :'MMULT' => 'mmult',
    :'MOD' => 'mod',
    :'NA' => 'na',
    :'NOT' => 'excel_not',
    :'NPV' => 'npv',
    :'NUMBER_OR_ZERO' => 'number_or_zero',
    :'OR' => 'excel_or',
    :'PI' => 'pi',
    :'PMT' => 'pmt',
    :'POWER' => 'power',
    :'PV' => 'pv',
    :'RANK' => 'rank',
    :'RIGHT' => 'right',
    :'ROUND' => 'round',
    :'ROUNDDOWN' => 'rounddown',
    :'ROUNDUP' => 'roundup',
    :'SUBSTITUTE' => 'substitute',
    :'SUBTOTAL' => 'subtotal',
    :'SUM' => 'sum',
    :'SUMIF' => 'sumif',
    :'SUMIFS' => 'sumifs',
    :'SUMPRODUCT' => 'sumproduct',
    :'TEXT' => 'text',
    :'TRIM' => 'trim',
    :'VALUE' => 'value',
    :'VLOOKUP' => 'vlookup',
    :'^' => 'power'
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
    strings = strings.map do |s|
      s = [:string, ""] if s == [:inlined_blank]
      s = map(s)
    end

    "string_join(#{strings.join(',')})"
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
