class ExtractArrayFormulaForCell
  
  attr_accessor :row_offset, :column_offset
  
  def map(ast)
    case ast.first
    when :array; map_array(ast)
    when :function; map_function(ast)
    else return ast
    end
  end
  
  def map_array(ast)
    if (@row_offset + 1) >= ast.length 
      if ast.length == 2
        @row_offset = 0
      else
        return [:error, "#N/A"]
      end
    end

    if (@column_offset + 1) >= ast[1].length
      if ast[1].length == 2
        @column_offset = 0
      else
        return [:error, "#N/A"]
      end
    end
    
    ast[@row_offset+1][@column_offset+1] # plus ones to skip tthe [:array,[:row,"cell"]] symbols
  end
  
  FUNCTIONS_THAT_CAN_RETURN_ARRAYS = %w{INDEX MMULT}
  
  def map_function(ast)
    return ast unless FUNCTIONS_THAT_CAN_RETURN_ARRAYS.include?(ast[1])
    [:function, "INDEX", ast, [:number, (@row_offset+1).to_s], [:number, (column_offset+1).to_s]]
  end
  
end

class RewriteArrayFormulae
  def self.rewrite(*args)
    new.rewrite(*args)
  end
  
  def rewrite(input)
    @output = {}
    input.each do |ref, details|
      sheet = ref[0]
      array_range = details[0]
      ast = details[1]
      array_formula(sheet, array_range, ast)
    end
    @output
  end
  
  def array_formula(sheet, array_range, array_ast)
    array_range = "#{array_range}:#{array_range}" unless array_range.include?(':')
    array_range = Area.for(array_range)
    array_range.calculate_excel_variables
    start_reference = array_range.excel_start
    mapper = ExtractArrayFormulaForCell.new
    
    # Then we rewrite each of the subsidiaries
    array_range.offsets.each do |row,column|
      mapper.row_offset = row
      mapper.column_offset = column
      ref = start_reference.offset(row,column)
      @output[[sheet, ref]] = mapper.map(array_ast)
    end
  end
  
end
