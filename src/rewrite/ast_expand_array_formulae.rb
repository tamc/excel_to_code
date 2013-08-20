require_relative '../excel'

class AstExpandArrayFormulae
    
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
    left = map(left)
    right = map(right)
    return [:arithmetic, left, operator, right] unless array?(left,right)
    
    map_arrays([left,right]) do |arrayed|
      [:arithmetic,arrayed[0],operator,arrayed[1]]
    end
  end

  def comparison(left,operator,right)
    left = map(left)
    right = map(right)
    return [:comparison, left, operator, right] unless array?(left,right)

    map_arrays([left,right]) do |arrayed|
      [:comparison,arrayed[0],operator,arrayed[1]]
    end
  end
  
  def string_join(*strings)
    strings = strings.map { |s| map(s) }
    return [:string_join, *strings] unless array?(*strings)
    map_arrays(strings) do |arrayed_strings|
      [:string_join, *arrayed_strings]
    end
  end
  
  def map_arrays(arrays, &block)
    # Turn them into ruby arrays
    arrays = arrays.map { |a| array_ast_to_ruby_array(a) }

    # Find the largest one
    max_rows = arrays.max { |a,b| a.length <=> b.length }.length
    max_columns = arrays.max { |a,b| a.first.length <=> b.first.length }.first.length
    
    # Convert any single rows into an array of the right size
    arrays = arrays.map { |a| a.length == 1 ? Array.new(max_rows,a.first) : a }
    
    # Convert any single columns into an array of the right size
    arrays = arrays.map { |a| a.first.length == 1 ? Array.new(max_columns,a.flatten(1)).transpose : a }
    
    # Now iterate through
    return [:array, *max_rows.times.map do |row|
      [:row, *max_columns.times.map do |column| 
        block.call(arrays.map do |a|
          a[row][column] || [:error, "#N/A"]
        end)
      end]
    end]
  end
  
  FUNCTIONS_THAT_ACCEPT_RANGES_FOR_ALL_ARGUMENTS = %w{AVERAGE COUNT COUNTA MAX MIN SUM SUMPRODUCT MMULT}
  
  def function(name,*arguments)
    if FUNCTIONS_THAT_ACCEPT_RANGES_FOR_ALL_ARGUMENTS.include?(name)
      [:function, name, *arguments.map { |a| map(a) }]
    elsif respond_to?("map_#{name.downcase}")
      send("map_#{name.downcase}",*arguments)
    else
      function_that_does_not_accept_ranges(name,arguments)
    end
  end
  
  def function_that_does_not_accept_ranges(name,arguments)
    return [:function, name] if arguments.empty?
    array_map arguments, name, *Array.new(arguments.length,false)
  end
  
  def map_match(*args)
    a = array_map args, 'MATCH', false, true, false
    a
  end
  
  def map_subtotal(*args)
    array_map args, 'SUBTOTAL', false, *Array.new(args.length-1,true)
  end
  
  def map_index(*args)
    array_map args, 'INDEX', true, false, false
  end
  
  def map_sumif(*args)
    array_map args, 'SUMIF', true, false, true
  end
  
  def map_sumifs(*args)
    if args.length > 3
      array_map args, 'SUMIFS', true, true, false, *([true,false]*((args.length-3)/2))
    else
      array_map args, 'SUMIFS', true, true, false
    end
  end
  
  def map_vlookup(*args)
    array_map args, "VLOOKUP", false, true, false, false
  end
  
  private
  
  def no_need_to_array?(args, ok_to_be_an_array)
    ok_to_be_an_array.each_with_index do |array_ok,i|
      next if array_ok
      break unless args[i]
      return false if args[i].first == :array
    end
    true
  end
  
  def array_map(args,function,*ok_to_be_an_array)
    args = args.map { |a| map(a) }
    return [:function, function, *args ] if no_need_to_array?(args,ok_to_be_an_array)

    # Turn the relevant arguments into ruby arrays and store the dimensions
    # Enumerable#max and Enumerable#min don't return Enumerators, so can't do it using those methods
    max_rows = 1
    max_columns = 1
    args = args.map.with_index do |a,i| 
      unless ok_to_be_an_array[i]
        a = array_ast_to_ruby_array(a)
        r = a.length
        c = a.first.length
        max_rows = r if r > max_rows
        max_columns = c if c > max_columns
      end
      a
    end
        
    # Convert any single rows into an array of the right size
    args = args.map.with_index { |a,i| (!ok_to_be_an_array[i] && a.length == 1) ? Array.new(max_rows,a.first) : a }
    
    # Convert any single columns into an array of the right size
    args = args.map.with_index { |a,i| (!ok_to_be_an_array[i] && a.first.length == 1) ? Array.new(max_columns,a.flatten(1)).transpose : a }
    
    # Now iterate through
    return [:array, *max_rows.times.map do |row|
      [:row, *max_columns.times.map do |column| 
        [:function, function, *args.map.with_index do |a,i|
          if ok_to_be_an_array[i]
            a
          else
            a[row][column] || [:error, "#N/A"]
          end
        end]
      end]
    end]
  end
  
  def array?(*args)
    args.any? { |a| a.first == :array }
  end
  
  def array_ast_to_ruby_array(array_ast)
    return [[array_ast]] unless array_ast.first == :array
    array_ast[1..-1].map do |row_ast|
      row_ast[1..-1].map do |cell|
        cell
      end
    end
  end
  
end
