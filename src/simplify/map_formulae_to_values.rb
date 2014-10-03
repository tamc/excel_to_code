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

  DO_NOT_MAP = {:number => true, :string => true, :blank => true, :inlined_blank => true, :null => true, :error => true, :boolean_true => true, :boolean_false => true, :sheet_reference => true, :cell => true}

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
    l = @calculator.number_argument(value(left))
    r = @calculator.number_argument(value(right))
    if (l == :not_a_value) && (r == :not_a_value)
      return ast
    elsif (l != :not_a_value) && (r != :not_a_value)
      ast.replace(formula_value(operator.last,l,r))
    # SPECIAL CASES
    elsif l == 0
      case operator.last
      when :+ 
        ast.replace(n(right))
      when :*
        ast.replace([:number, 0])
      end
    elsif r == 0
      case operator.last
      when :+, :-
        ast.replace(n(left))
      when :* 
        ast.replace([:number, 0])
      when :/
        ast.replace([:error, :'#DIV/0!'])
      when :^
        ast.replace([:number, 1])
      end
    elsif l == 1
      case operator.last
      when :*
        ast.replace(n(right))
      when :^
        ast.replace([:number, 1])
      end
    elsif r == 1
      case operator.last
      when :*, :/, :^
        ast.replace(n(left))
      end
    end
    ast
  end

  def n(ast)
    return ast if ast[0] == :function && ast[1] == :ENSURE_IS_NUMBER
    [:function, :ENSURE_IS_NUMBER, ast]
  end

  def comparison(ast)
    left, operator, right = ast[1], ast[2], ast[3]
    l = value(left)
    r = value(right)
    return ast if (l == :not_a_value) || (r == :not_a_value)
    ast.replace(formula_value(operator.last,l,r))
  end

  # [:percentage, number]
  def percentage(ast)
    ast.replace(ast_for_value(value([:percentage, ast[1]])))
  end
  
  # [:string_join, stringA, stringB, ...]
  def string_join(ast)
    values = ast[1..-1].map do |a| 
        value(a, "")
    end
    return if values.any? { |a| a == :not_a_value }
    ast.replace(ast_for_value(@calculator.string_join(*values)))
  end
  
  # [:function, function_name, arg1, arg2, ...]
  def function(ast)
    name = ast[1]
    return if name == :INDIRECT
    return if name == :OFFSET
    return if name == :COLUMN
    return if name == :ROW
    if respond_to?("map_#{name.to_s.downcase}")
      send("map_#{name.to_s.downcase}",ast)
    else
      normal_function(ast)
    end
  end

  def normal_function(ast, inlined_blank = 0)
    values = ast[2..-1].map { |a| value(a, inlined_blank) }
    return if values.any? { |a| a == :not_a_value }
    ast.replace(formula_value( ast[1],*values))
  end

  def map_text(ast)
    values = ast[2..-1].map { |a| value(a, nil) }
    return if values.any? { |a| a == :not_a_value }
    ast.replace(formula_value( ast[1],*values))
  end

  def map_if(ast)
    condition_ast = ast[2]
    true_option_ast = ast[3]
    false_option_ast = ast[4] || [:boolean_false]

    condition_value = value(condition_ast)
    return if condition_value == :not_a_value

    case condition_value
    when Symbol
      ast.replace(condition_ast)
    when String
      ast.replace([:error, :"#VALUE!"])
    when false, 0
      ast.replace(false_option_ast)
    else
      ast.replace(true_option_ast)
    end
    ast
  end

  def map_right(ast)
    normal_function(ast, "")
  end

  def map_left(ast)
    normal_function(ast, "")
  end

  def map_mid(ast)
    normal_function(ast, "")
  end

  def map_len(ast)
    normal_function(ast, "")
  end

  def map_find(ast)
    normal_function(ast, "")
  end

  def map_isblank(ast)
    normal_function(ast,nil)
  end

  OK_CHECK_RANGE_TYPES = [:sheet_reference, :cell, :area, :array, :number, :string, :boolean_true, :boolean_false]

  def map_sumifs(ast)
    values = ast[3..-1].map.with_index { |a,i| value(a, (i % 2) == 0 ? 0 : nil ) }
    return if values.all? { |a| a == :not_a_value } # Nothing to be done
    return attempt_to_reduce_sumifs(ast) if values.any? { |a| a == :not_a_value } # Maybe a reduction to be done
    return unless OK_CHECK_RANGE_TYPES.include?(ast[2].first)
    sum_value = value(ast[2])
    if sum_value == :not_a_value # i.e., a sheet_reference, :cell or :area
      partially_map_sumifs(ast)
    else
      ast.replace(formula_value( ast[1], sum_value, *values))
    end
  end

  def partially_map_sumifs(ast)
    values = ast[3..-1].map.with_index { |a,i| value(a, (i % 2) == 0 ? 0 : nil ) }
    sum_range = array_as_values(ast[2]).flatten(1)
    indexes = @calculator._filtered_range_indexes(sum_range, *values)
    if indexes.is_a?(Symbol)
      new_ast = ast_for_value(indexes)
    elsif indexes.empty?
      new_ast = [:number, 0]
    else
      new_ast = [:function, :SUM, *sum_range.values_at(*indexes)]
    end
    if new_ast != ast
      @replacements_made_in_the_last_pass += 1
      ast.replace(new_ast)
    end
  end

  # FIXME: Ends up making everything single column. Is that ok?!
  def attempt_to_reduce_sumifs(ast)
    return unless OK_CHECK_RANGE_TYPES.include?(ast[2].first)
    # First combine into a series of checks
    criteria_that_can_be_resolved = []
    criteria_that_cant_be_resolved = []
    ast[3..-1].each_slice(2) do |check|
      # Give up unless we have something that can actually be used
      return unless OK_CHECK_RANGE_TYPES.include?(check[0].first)
      check_range_value = value(check[0])
      check_criteria_value = value(check[1])
      if check_range_value == :not_a_value || check_criteria_value == :not_a_value
        criteria_that_cant_be_resolved << check
      else
        criteria_that_can_be_resolved << [check_range_value, check_criteria_value]
      end
    end
    return if criteria_that_can_be_resolved.empty?
    sum_range = array_as_values(ast[2]).flatten(1)
    indexes = @calculator._filtered_range_indexes(sum_range, *criteria_that_can_be_resolved.flatten(1))
    if indexes.is_a?(Symbol)
      return
    elsif indexes.empty?
      new_ast = [:number, 0]
    else
      new_ast = [:function, :SUMIFS]
      new_ast << ast_for_array(sum_range.values_at(*indexes))
      criteria_that_cant_be_resolved.each do |check|
        new_ast << ast_for_array(array_as_values(check.first).flatten(1).values_at(*indexes))
        new_ast << check.last
      end
    end
    if new_ast != ast
      @replacements_made_in_the_last_pass += 1
      ast.replace(new_ast)
    end
  end


  # [:function, "COUNT", range]
  def map_count(ast)
    values = ast[2..-1].map { |a| value(a, nil) }
    return if values.any? { |a| a == :not_a_value }
    ast.replace(formula_value( ast[1],*values))
  end
  
  # [:function, "INDEX", array, row_number, column_number]
  def map_index(ast)
    return map_index_with_only_two_arguments(ast) if ast.length == 4

    array_mapped = ast[2] 
    return ast.replace(ast[2]) if ast[2].first == :error
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
    return ast.replace(ast[2]) if ast[2].first == :error
    row_as_number = value(ast[3])
    return if row_as_number == :not_a_value
    array_as_values = array_as_values(array_mapped)
    return unless array_as_values
    result = @calculator.send(MapFormulaeToRuby::FUNCTIONS[:INDEX],array_as_values,row_as_number)
    result = [:number, 0] if result == [:blank]
    result = ast_for_value(result)
    ast.replace(result)
  end

  # [:function, :SUM, a, b, c...] 
  def map_sum(ast)
      values = ast[2..-1].map { |a| value(a) }
      return partially_map_sum(ast) if values.any? { |a| a == :not_a_value }
      ast.replace(formula_value(:SUM,*values))
  end

  def partially_map_sum(ast)
    number_total = 0
    not_number_array = []
    ast[2..-1].each do |a|
      result = filter_numbers_and_not(a)
      number_total += result.first
      not_number_array.concat(result.last)
    end
    if number_total == 0 && not_number_array.empty?
      new_ast = [:number, number_total]
      if new_ast != ast
        @replacements_made_in_the_last_pass += 1
        ast.replace(new_ast)
      end
    # FIXME: Will I be haunted by this? What if doing a sum of something that isn't a number
    # and so what is expected is a VALUE error?. YES. This doesn't work well.
    elsif ast.length == 3 && [:cell, :sheet_reference].include?(ast[2].first)
      new_ast = n(ast[2])
      if new_ast != ast
        @replacements_made_in_the_last_pass += 1
        ast.replace(new_ast)
      end
    elsif ast.length == 3 && ast[2][0] == :function && ast[2][1] == :ENSURE_IS_NUMBER
      new_ast = ast[2]
      if new_ast != ast
        @replacements_made_in_the_last_pass += 1
        ast.replace(new_ast)
      end
    else
      new_ast = [:function, :SUM].concat(not_number_array)
      new_ast.push([:number, number_total]) unless number_total == 0
      if new_ast != ast
        @replacements_made_in_the_last_pass += 1
        ast.replace(new_ast)
      end
    end
    ast
  end

  def filter_numbers_and_not(ast)
    number_total = 0
    not_number_array = []
    case ast.first
    when :array
      array_total = 0
      array_not_numbers = []
      # First we just have a go at splitting the array into a list of numbers
      # and not numbers.
      array_as_values(ast).each do |row|
        row.each do |c|
          result = filter_numbers_and_not(c)
          array_total += result.first
          array_not_numbers.concat(result.last)
        end
      end
      # If there are no not_numbers, or only one, we are good
      if array_not_numbers.length <= 1
        number_total += array_total
        not_number_array.concat(array_not_numbers)
      # If there are more than on not_numbers we aren't neccessarily good
      # unless all those not numbers are simple
      elsif array_not_numbers.all? { |c| [:cell, :area, :sheet_reference].include?(c.first)}
        number_total += array_total
        not_number_array.concat(array_not_numbers)
      # Otherwise, leave that array alone
      else
        not_number_array.push(ast)
      end
    when :blank, :number, :percentage, :string, :boolean_true, :boolean_false
      number = @calculator.number_argument(value(ast))
      if number.is_a?(Symbol)
        not_number_array.push(ast)
      else
        number_total += number
      end
    else
      not_number_array.push(ast)
    end
    [number_total, not_number_array]
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

  # FIXME: Assumes single column. Not wise?
  def ast_for_array(array)
    [:array,*array.map { |row| [:row, row ]}]
  end

  ERRORS = {
    :"#NAME?" => :name,
    :"#VALUE!" => :value,
    :"#DIV/0!" => :div0,
    :"#REF!" => :ref,
    :"#N/A" => :na,
    :"#NUM!" => :num
  }
    
  def value(ast, inlined_blank = 0)
    return extract_values_from_array(ast, inlined_blank) if ast.first == :array
    case ast.first
    when :blank; nil
    when :inlined_blank; inlined_blank
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
  
  def extract_values_from_array(ast, inlined_blank = 0)
    ast[1..-1].map do |row|
      row[1..-1].map do |cell|
        v = value(cell, inlined_blank)
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
