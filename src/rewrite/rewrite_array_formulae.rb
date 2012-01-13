require_relative 'rewrite_array_formulae_ast'

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
    mapper = RewriteArrayFormulaAst.new
    
    # Then we rewrite each of the subsidiaries
    array_range.offsets.each do |row,column|
      mapper.row_offset = row
      mapper.column_offset = column
      ref = start_reference.offset(row,column)
      output.puts "#{ref}\t#{mapper.map(array_ast).inspect}"
    end
  end
  
end