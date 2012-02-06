class ExtractArrayFormulaForCell
  
  attr_accessor :row_offset, :column_offset
  
  def map(ast)
    return ast unless ast.first == :array

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
  
end

class RewriteArrayFormulae
  def self.rewrite(input,output)
    new.rewrite(input,output)
  end
  
  def rewrite(input,output)
    input.lines do |line|
      ref, array_range, formula = line.split("\t")
      array_formula(formula,array_range,output)
    end
  end
  
  def array_formula(formula,array_range,output)
    array_ast = eval(formula)
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
      output.puts "#{ref}\t#{mapper.map(array_ast).inspect}"
    end
  end
  
end