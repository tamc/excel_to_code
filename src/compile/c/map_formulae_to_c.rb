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
    '<' => 'less_than',
    '<=' => 'less_than_or_equal',
    '<>' => 'not_equal',
    '=' => 'excel_equal',
    '>' => 'more_than',
    '>=' => 'more_than_or_equal',
    'ABS' => 'excel_abs',
    'AND' => 'excel_and',
    'AVERAGE' => 'average',
    'CHOOSE' => 'choose',
    'CONCATENATE' => 'string_join',
    'COSH' => 'cosh',
    'COUNT' => 'count',
    'COUNTA' => 'counta',
    'FIND2' => 'find_2',
    'FIND3' => 'find',
    'HLOOKUP3' => 'hlookup_3',
    'HLOOKUP4' => 'hlookup',
    'IF2' => 'excel_if_2',
    'IF3' => 'excel_if',
    'IFERROR' => 'iferror',
    'INDEX2' => 'excel_index_2',
    'INDEX3' => 'excel_index',
    'INT' => 'excel_int',
    'ISNUMBER' => 'excel_isnumber',
    'LARGE' => 'large',
    'LEFT1' => 'left_1',
    'LEFT2' => 'left',
    'LOG1' => 'excel_log',
    'LOG2' => 'excel_log_2',
    'MATCH2' => 'excel_match_2',
    'MATCH3' => 'excel_match',
    'MAX' => 'max',
    'MIN' => 'min',
    'MMULT' => 'mmult',
    'MOD' => 'mod',
    'PMT' => 'pmt',
    'PV3' => 'pv_3',
    'PV4' => 'pv_4',
    'PV5' => 'pv_5',
    'RANK2' => 'rank_2',
    'RANK3' => 'rank',
    'ROUND' => 'excel_round',
    'ROUNDDOWN' => 'rounddown',
    'ROUNDUP' => 'roundup',
    'string_join' => 'string_join',
    'SUBTOTAL' => 'subtotal',
    'SUM' => 'sum',
    'SUMIF2' => 'sumif_2',
    'SUMIF3' => 'sumif',
    'SUMIFS' => 'sumifs',
    'TEXT2' => 'text',
    'SUMPRODUCT' => 'sumproduct',
    'VLOOKUP3' => 'vlookup_3',
    'VLOOKUP4' => 'vlookup',
    '^' => 'power',
    'POWER' => 'power'
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
    if self.respond_to?("function_#{function_name.downcase}")
      send("function_#{function_name.downcase}",*arguments)
    # Some arguments can take any number of arguments, which we need to treat separately
    elsif FUNCTIONS_WITH_ANY_NUMBER_OF_ARGUMENTS.include?(function_name)
      any_number_of_argument_function(function_name,arguments)

    # Check for whether this function has variants based on the number of arguments
    elsif FUNCTIONS.has_key?("#{function_name}#{arguments.size}")
      "#{FUNCTIONS["#{function_name}#{arguments.size}"]}(#{arguments.map { |a| map(a) }.join(",")})"

    # Then check for whether it is just a standard type
    elsif FUNCTIONS.has_key?(function_name)
      "#{FUNCTIONS[function_name]}(#{arguments.map { |a| map(a) }.join(",")})"

    else
      raise NotSupportedException.new("Function #{function_name} with #{arguments.size} arguments not supported")
    end
  end
  
  FUNCTIONS_WITH_ANY_NUMBER_OF_ARGUMENTS = %w{SUM AND AVERAGE COUNT COUNTA MAX MIN SUMPRODUCT CONCATENATE}
  
  def function_pi() 
    "M_PI"
  end
  
  def function_choose(index,*arguments)
    "#{FUNCTIONS["CHOOSE"]}(#{map(index)}, #{map_arguments_to_array(arguments)})"
  end
  
  def function_subtotal(type,*arguments)
    "#{FUNCTIONS["SUBTOTAL"]}(#{map(type)}, #{map_arguments_to_array(arguments)})"
  end

  def function_sumifs(sum_range,*criteria)
    "#{FUNCTIONS["SUMIFS"]}(#{map(sum_range)}, #{map_arguments_to_array(criteria)})"
  end

  
  def any_number_of_argument_function(function_name,arguments)    
    "#{FUNCTIONS[function_name]}(#{map_arguments_to_array(arguments)})"
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
      "_#{reference}()"
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
    initializers << "ExcelValue #{range_name} = new_excel_range(#{array_name},#{number_of_rows},#{number_of_columns});"

    range_name
  end


end
