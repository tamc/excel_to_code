require_relative '../excel'

class AstExpandArrayFormulae
    
  def map(ast)
    return ast unless ast.is_a?(Array)
    operator = ast[0]
    send(operator, ast) if respond_to?(operator)
    ast.each {|a| map(a) }
    ast
  end

  # Format [:arithmetic, left, operator, right]
  def arithmetic(ast)
    ast.each {|a| map(a) }
    return unless array?(ast[1], ast[3])
    
    ast.replace(
      map_arrays([ast[1],ast[3]]) do |arrayed|
        [:arithmetic,arrayed[0],ast[2],arrayed[1]]
      end
    )
  end

  # Format [:comparison, left, operator, right]
  def comparison(ast)
    ast.each {|a| map(a) }
    return unless array?(ast[1], ast[3])
    
    ast.replace(
      map_arrays([ast[1],ast[3]]) do |arrayed|
        [:comparison,arrayed[0],ast[2],arrayed[1]]
      end
    )
  end
  
  # Format [:string_join, stringA, stringB, ...]
  def string_join(ast)
    ast.each {|a| map(a) }
    return unless array?(*ast[1..-1])
    ast.replace(
      map_arrays(ast[1..-1]) do |arrayed_strings|
        [:string_join, *arrayed_strings]
      end
    )
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
          (a[row] && a[row][column]) || CachingFormulaParser.map([:error, :"#N/A"])
        end)
      end]
    end]
  end
  
  FUNCTIONS_THAT_ACCEPT_RANGES_FOR_ALL_ARGUMENTS = {'AVERAGE' => true, 'COUNT' => true, 'COUNTA' => true, 'MAX' => true, 'MIN' => true, 'SUM' => true, 'SUMPRODUCT' => true, 'MMULT' => true}
  
  # Format [:function, function_name, arg1, arg2, ...]
  def function(ast)
    name = ast[1]
    arguments = ast[2..-1]
    if FUNCTIONS_THAT_ACCEPT_RANGES_FOR_ALL_ARGUMENTS.has_key?(name)
      ast.each { |a| map(a) }
      return # No need to alter anything
    elsif respond_to?("map_#{name.downcase}")
      # These typically have some arguments that accept ranges, but not all
      send("map_#{name.downcase}",ast)
    else
      function_that_does_not_accept_ranges(ast)
    end
  end
  
  def function_that_does_not_accept_ranges(ast)
    return if ast.length == 2
    name = ast[1]
    arguments = ast[2..-1]
    array_map(ast, *Array.new(arguments.length,false))
  end
  
  def map_match(ast)
    array_map(ast, false, true, false)
  end
  
  def map_subtotal(ast)
    array_map ast, false, *Array.new(ast.length-3,true)
  end
  
  def map_index(ast)
    array_map ast, true, false, false
  end
  
  def map_sumif(ast)
    array_map ast, true, false, true
  end
  
  def map_sumifs(ast)
    if ast.length > 5
      array_map ast, true, true, false, *([true,false]*((ast.length-5)/2))
    else
      array_map ast, true, true, false
    end
  end
  
  def map_vlookup(ast)
    array_map ast, false, true, false, false
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
  
  def array_map(ast,*ok_to_be_an_array)
    ast.each { |a| map(a) }

    return if no_need_to_array?(ast[2..-1],ok_to_be_an_array)

    # Turn the relevant arguments into ruby arrays and store the dimensions
    # Enumerable#max and Enumerable#min don't return Enumerators, so can't do it using those methods
    max_rows = 1
    max_columns = 1
    args = ast[2..-1].map.with_index do |a,i| 
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
    ast.replace( [:array, *max_rows.times.map do |row|
      [:row, *max_columns.times.map do |column| 
        [:function, ast[1], *args.map.with_index do |a,i|
          if ok_to_be_an_array[i]
            a
          else
            a[row][column] || [:error, :"#N/A"]
          end
        end]
      end]
    end])
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
