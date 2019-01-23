require_relative 'map_values_to_c'

class MapFormulaeToC < MapValuesToC

  attr_accessor :sheet_names
  attr_accessor :worksheet
  attr_reader :initializers
  attr_reader :counter
  attr_accessor :allow_unknown_functions

  def initialize
    reset
  end

  def reset
    @initializers = []
    @counter = 0
  end

  FUNCTIONS = {
    :'*' => 'multiply',
    :'+' => 'add',
    :'-' => 'subtract',
    :'/' => 'divide',
    :'<' => 'less_than',
    :'<=' => 'less_than_or_equal',
    :'<>' => 'not_equal',
    :'=' => 'excel_equal',
    :'>' => 'more_than',
    :'>=' => 'more_than_or_equal',
    :'ABS' => 'excel_abs',
    :'AND' => 'excel_and',
    :'AVERAGE' => 'average',
    :'AVERAGEIFS' => 'averageifs',
    :'CEILING' => 'ceiling_2',
    :'_xlfn.CEILING.MATH2' => 'excel_ceiling_math_2',
    :'_xlfn.CEILING.MATH3' => 'excel_ceiling_math',
    :'CHAR' => 'excel_char',
    :'CHOOSE' => 'choose',
    :'CONCATENATE' => 'string_join',
    :'COSH' => 'cosh',
    :'COUNT' => 'count',
    :'COUNTA' => 'counta',
    :'COUNTIFS' => 'countifs',
    :'ENSURE_IS_NUMBER' => 'ensure_is_number',
    :'EXP' => 'excel_exp',
    :'FIND2' => 'find_2',
    :'FIND3' => 'find',
    :'FLOOR' => 'excel_floor',
    :'FORECAST' => 'forecast',
    :'_xlfn.FORECAST.LINEAR' => 'forecast',
    :'HLOOKUP3' => 'hlookup_3',
    :'HLOOKUP4' => 'hlookup',
    :'IF2' => 'excel_if_2',
    :'IF3' => 'excel_if',
    :'IFERROR' => 'iferror',
    :'ISERR' => 'iserr',
    :'ISERROR' => 'iserror',
    :'INDEX2' => 'excel_index_2',
    :'INDEX3' => 'excel_index',
    :'INT' => 'excel_int',
    :'ISNUMBER' => 'excel_isnumber',
    :'ISBLANK' => 'excel_isblank',
    :'LARGE' => 'large',
    :'LEFT1' => 'left_1',
    :'LEFT2' => 'left',
    :'LEN' => 'len',
    :'LN' => 'ln',
    :'LOG10' => 'excel_log',
    :'LOG1' => 'excel_log',
    :'LOG2' => 'excel_log_2',
    :'MATCH2' => 'excel_match_2',
    :'MATCH3' => 'excel_match',
    :'MAX' => 'max',
    :'MIN' => 'min',
    :'MMULT' => 'mmult',
    :'MOD' => 'mod',
    :'MROUND' => 'mround',
    :'NA' => 'na',
    :'NOT' => 'excel_not',
    :'OR' => 'excel_or',
    :'NPV' => 'npv',
    :'NUMBER_OR_ZERO' => 'number_or_zero',
    :'PMT3' => 'pmt',
    :'PMT4' => 'pmt_4',
    :'PMT5' => 'pmt_5',
    :'PRODUCT' => 'product',
    :'PV3' => 'pv_3',
    :'PV4' => 'pv_4',
    :'PV5' => 'pv_5',
    :'RATE' => 'rate',
    :'RANK2' => 'rank_2',
    :'RANK3' => 'rank',
    :'RIGHT1' => 'right_1',
    :'RIGHT2' => 'right',
    :'ROUND' => 'excel_round',
    :'ROUNDDOWN' => 'rounddown',
    :'ROUNDUP' => 'roundup',
    :'string_join' => 'string_join',
    :'SUBTOTAL' => 'subtotal',
    :'SUBSTITUTE3' => 'substitute_3',
    :'SUBSTITUTE4' => 'substitute_4',
    :'SUM' => 'sum',
    :'SUMIF2' => 'sumif_2',
    :'SUMIF3' => 'sumif',
    :'SUMIFS' => 'sumifs',
    :'TEXT2' => 'text',
    :'SUMPRODUCT' => 'sumproduct',
    :'VALUE' => 'value',
    :'VLOOKUP3' => 'vlookup_3',
    :'VLOOKUP4' => 'vlookup',
    :'^' => 'power',
    :'POWER' => 'power',
    :'SQRT' => 'excel_sqrt',
    :'curve5' => 'curve_5',
    :'curve' => 'curve',
    :'scurve4' => 'scurve_4',
    :'scurve' => 'scurve',
    :'halfscurve4' => 'halfscurve_4',
    :'halfscurve' => 'halfscurve',
    :'lcurve4' => 'lcurve_4',
    :'lcurve' => 'lcurve',
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
    any_number_of_argument_function('string_join',strings)
  end

  def comparison(left,operator,right)
    "#{FUNCTIONS[operator.last]}(#{map(left)},#{map(right)})"
  end

  def function(function_name,*arguments)
    # Some functions are special cases
    if self.respond_to?("function_#{function_name.to_s.downcase}")
      send("function_#{function_name.to_s.downcase}",*arguments)
    # Some arguments can take any number of arguments, which we need to treat separately
    elsif FUNCTIONS_WITH_ANY_NUMBER_OF_ARGUMENTS.include?(function_name.to_s)
      any_number_of_argument_function(function_name,arguments)

    # Check for whether this function has variants based on the number of arguments
    elsif FUNCTIONS.has_key?("#{function_name.to_s}#{arguments.size}".to_sym)
      "#{FUNCTIONS["#{function_name.to_s}#{arguments.size}".to_sym]}(#{arguments.map { |a| map(a) }.join(",")})"

    # Then check for whether it is just a standard type
    elsif FUNCTIONS.has_key?(function_name.to_sym)
      "#{FUNCTIONS[function_name.to_sym]}(#{arguments.map { |a| map(a) }.join(",")})"

    # Optionally, can dump unknown functions
  elsif self.allow_unknown_functions
      "#{function_name.to_s.downcase}(#{arguments.map { |a| map(a) }.join(",")})"

    # But default is to raise an error
    else
      raise NotSupportedException.new("Function #{function_name} with #{arguments.size} arguments not supported")
    end
  end

  FUNCTIONS_WITH_ANY_NUMBER_OF_ARGUMENTS = %w{
    SUM 
    PRODUCT 
    AND 
    OR
    AVERAGE 
    COUNT 
    COUNTA 
    MAX 
    MIN 
    SUMPRODUCT 
    COUNTIFS 
    CONCATENATE
  }

  def function_pi()
    "M_PI"
  end

  def function_choose(index,*arguments)
    "#{FUNCTIONS[:CHOOSE]}(#{map(index)}, #{map_arguments_to_array(arguments)})"
  end

  def function_subtotal(type,*arguments)
    "#{FUNCTIONS[:SUBTOTAL]}(#{map(type)}, #{map_arguments_to_array(arguments)})"
  end

  def function_sumifs(sum_range,*criteria)
    "#{FUNCTIONS[:SUMIFS]}(#{map(sum_range)}, #{map_arguments_to_array(criteria)})"
  end

  def function_averageifs(average_range,*criteria)
    "#{FUNCTIONS[:AVERAGEIFS]}(#{map(average_range)}, #{map_arguments_to_array(criteria)})"
  end

  def function_npv(rate,*cash_flows)
    "#{FUNCTIONS[:NPV]}(#{map(rate)}, #{map_arguments_to_array(cash_flows)})"
  end

  def function_if(condition, true_case, false_case = [:boolean_false])
    true_code = map(true_case)
    false_code = map(false_case)

    condition_name = "condition#{@counter}"
    result_name = "ifresult#{@counter}"
    @counter += 1

    initializers << "ExcelValue #{condition_name} = #{map(condition)};"
    initializers << "ExcelValue #{result_name};"
    initializers << "switch(#{condition_name}.type) {"
    initializers << "case ExcelBoolean:"
  	initializers << "  if(#{condition_name}.number == true) {"
    initializers << "    #{result_name} = #{true_code};"
    initializers << "  } else {"
    initializers << "    #{result_name} = #{false_code};"
    initializers << "  }"
    initializers << "  break;"
  	initializers << "case ExcelNumber:"
    initializers << "  if(#{condition_name}.number == 0) {"
    initializers << "    #{result_name} = #{false_code};"
    initializers << "  } else {"
    initializers << "    #{result_name} = #{true_code};"
    initializers << "  }"
    initializers << "  break;"
	  initializers << "case ExcelEmpty: "
    initializers << "  #{result_name} = #{false_code};"
    initializers << "  break;"
	  initializers << "case ExcelString:"
    initializers << "case ExcelRange:"
    initializers << "  #{result_name} = VALUE;"
    initializers << "  break;"
  	initializers << "case ExcelError:"
    initializers << "  #{result_name} = #{condition_name};"
    initializers << "  break;"
    initializers << "}"

    return result_name
  end

  def any_number_of_argument_function(function_name,arguments)
    "#{FUNCTIONS[function_name.to_sym]}(#{map_arguments_to_array(arguments)})"
  end

  def map_arguments_to_array(arguments)
    # First we have to create an excel array
    array_name = "array#{@counter}"
    @counter +=1
    arguments_size = arguments.size
    arguments = arguments.map { |a| map(a) }.join(',')
    initializers << "ExcelValue #{array_name}[] = {#{arguments}};"
    "#{arguments_size}, #{array_name}"
  end

  def cell(reference)
    # FIXME: What a cludge.
    if reference =~ /common\d+/
      "#{reference}()"
    else
      reference.to_s.downcase.gsub('$','')
    end
  end

  def sheet_reference(sheet,reference)
    "#{sheet_names[sheet]}_#{map(reference).to_s.downcase}()"
  end

  def array(*rows)
    # Make sure we get the right dimensions
    number_of_rows = rows.size
    number_of_columns = rows[0].size - 1
    size = number_of_rows * number_of_columns

    # First we have to create an excel array
    array_name = "array#{@counter}"
    @counter +=1

    initializers << "static ExcelValue #{array_name}[#{size}];"
    i = 0
    rows.each do |row|
      row.each do |cell|
        next if cell.is_a?(Symbol)
        initializers << "#{array_name}[#{i}] = #{map(cell)};"
        i += 1
      end
    end

    # Then we need to assign it to an excel value
    range_name = array_name+"_ev"
    initializers << "ExcelValue #{range_name} = EXCEL_RANGE(#{array_name},#{number_of_rows},#{number_of_columns});"

    range_name
  end


end
