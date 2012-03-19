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
    'COSH' => 'cosh',
    'COUNT' => 'count',
    'COUNTA' => 'counta',
    'FIND2' => 'find_2',
    'FIND3' => 'find',
    'IF2' => 'excel_if_2',
    'IF3' => 'excel_if',
    'IFERROR' => 'iferror',
    'INDEX' => 'index',
    'LEFT' => 'left',
    'MATCH2' => 'excel_match_2',
    'MATCH3' => 'excel_match',
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
  
  FUNCTIONS_WITH_ANY_NUMBER_OF_ARGUMENTS = %w{SUM AND AVERAGE COUNT COUNTA}
  
  def function_choose(index,*arguments)
    "#{FUNCTIONS["CHOOSE"]}(#{map(index)}, #{map_arguments_to_array(arguments)})"
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
      reference.downcase.gsub('$','')
    end
  end
  
  def sheet_reference(sheet,reference)
    "#{sheet_names[sheet]}_#{map(reference).downcase}()"
  end

  def array(*rows)
    # Make sure we get the right dimensions
    number_of_rows = rows.size
    number_of_columns = rows.max { |r| r.size }.size - 1
    
    # First we have to create an excel array
    array_name = "array#{@counter}"
    @counter +=1

    cells = rows.map do |r|
      r.shift if r.first == :row
      r.map do |c| 
        map(c)
      end
    end.flatten.join(',')
    
    initializers << "ExcelValue #{array_name}[] = {#{cells}};"
    
    # Then we need to assign it to an excel value
    range_name = array_name+"_ev"
    initializers << "ExcelValue #{range_name} = new_excel_range(#{array_name},#{number_of_rows},#{number_of_columns});"

    range_name
  end


end
