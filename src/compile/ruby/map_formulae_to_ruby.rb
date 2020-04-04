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
    :'CEILING' => 'ceiling',
    :'CELL' => 'cell',
    :'CHAR' => 'char',
    :'CHOOSE' => 'choose',
    :'CONCATENATE' => 'string_join',
    :'COSH' => 'cosh',
    :'COUNT' => 'count',
    :'COUNTA' => 'counta',
    :'COUNTIF' => 'countif',
    :'COUNTIFS' => 'countifs',
    :'ENSURE_IS_NUMBER' => 'ensure_is_number',
    :'EXACT' => 'exact',
    :'EXP' => 'exp',
    :'FILLGAPS' => 'fillgaps',
    :'FILLGAPS_IN_ARRAY' => 'fillgaps_in_array',
    :'FIND' => 'find',
    :'FLOOR' => 'floor',
    :'FORECAST' => 'forecast',
    :'HLOOKUP' => 'hlookup',
    :'HYPERLINK' => 'hyperlink',
    :'IF' => 'excel_if',
    :'IFERROR' => 'iferror',
    :'IFNA' => 'ifna',
    :'INDEX' => 'index',
    :'INT' => 'int',
    :'INTERPOLATE' => 'interpolate',
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
    :'LOOKUP' => 'lookup',
    :'LOWER' => 'lower',
    :'MATCH' => 'excel_match',
    :'MAX' => 'max',
    :'MID' => 'mid',
    :'MIN' => 'min',
    :'MMULT' => 'mmult',
    :'MOD' => 'mod',
    :'MROUND' => 'mround',
    :'NA' => 'na',
    :'NOT' => 'excel_not',
    :'NPV' => 'npv',
    :'NUMBER_OR_ZERO' => 'number_or_zero',
    :'OR' => 'excel_or',
    :'PI' => 'pi',
    :'PMT' => 'pmt',
    :'POWER' => 'power',
    :'PRODUCT' => 'product',
    :'PROJECT' => 'project',
    :'PROJECT_IN_ARRAY' => 'project_in_array',
    :'PV' => 'pv',
    :'RANK' => 'rank',
    :'RATE' => 'rate',
    :'REPLACE' => 'replace',
    :'RIGHT' => 'right',
    :'ROUND' => 'round',
    :'ROUNDDOWN' => 'rounddown',
    :'ROUNDUP' => 'roundup',
    :'SQRT' => 'sqrt',
    :'SUBSTITUTE' => 'substitute',
    :'SUBTOTAL' => 'subtotal',
    :'SUM' => 'sum',
    :'SUMIF' => 'sumif',
    :'SUMIFS' => 'sumifs',
    :'SUMPRODUCT' => 'sumproduct',
    :'TEXT' => 'text',
    :'TRIM' => 'trim',
    :'UNICODE' => 'unicode',
    :'VALUE' => 'value',
    :'VLOOKUP' => 'vlookup',
    :'^' => 'power',
    :'_xlfn.CEILING.MATH' => 'ceiling',
    :'_xlfn.FORECAST.LINEAR' => 'forecast',
    :'_xlfn.IFNA' => 'ifna',
    :'_xlfn.UNICODE' => 'unicode',
    :'curve' => 'curve',
    :'halfscurve' => 'halfscurve',
    :'lcurve' => 'lcurve',
    :'scurve' => 'scurve'
  }

  FUNCTIONS_THAT_CARE_ABOUT_BLANKS = {
    :'FILLGAPS_IN_ARRAY' => true,
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
      previous_thing_to_do_with_inline_blanks = @leave_inline_blank_as_nil
      @leave_inline_blank_as_nil = FUNCTIONS_THAT_CARE_ABOUT_BLANKS[function_name]
      result = "#{FUNCTIONS[function_name]}(#{arguments.map { |a| map(a) }.join(",")})"
      @leave_inline_blank_as_nil = previous_thing_to_do_with_inline_blanks
      result
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
