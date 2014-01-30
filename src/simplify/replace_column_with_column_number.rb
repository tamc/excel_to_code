class ReplaceColumnAndRowFunctionsAST
    
  attr_accessor :count_replaced
  attr_accessor :replacement_made
  attr_accessor :current_reference

  def initialize
    @count_replaced = 0
  end

  def replace(ast)
    @replacement_made = false
    map(ast)
    @replacement_made
  end
   
  def map(ast)
    return ast unless ast.is_a?(Array)
    function(ast) if ast[0] == :function
    ast.each { |a| map(a) }
    ast
  end
  
  # Should be of the form [:function, "COLUMN", [:sheet_reference, sheet, ref]] 

  def function(ast)
    return unless (ast[1] == :COLUMN || ast[1] == :ROW)
    if ast[2]
      if ast[2][0] == :cell
        reference = Reference.for(ast[2][1])
      elsif ast[2][0] == :sheet_reference
        reference = Reference.for(ast[2][2][1])
      else
        return
      end
    else
      reference = Reference.for(@current_reference)
    end 
    reference.calculate_excel_variables
    @count_replaced += 1
    @replacement_made = true
    if ast[1] == :COLUMN
      ast.replace( CachingFormulaParser.map([:number, reference.excel_column_number]))
    else
      ast.replace( CachingFormulaParser.map([:number, reference.excel_row_number]))
    end
  end

end
  

class ReplaceColumnAndRowFunctions
    
  def self.replace(*args)
    self.new.replace(*args)
  end

  attr_accessor :count_replaced
  
  def replace(input,output)
    rewriter = ReplaceColumnAndRowFunctionsAST.new
    input.each_line do |line|
      # Looks to match lines with references
      if line =~ /(:COLUMN|:ROW)/
        ref, ast = line.split("\t")
        rewriter.current_reference = ref
        output.puts "#{ref}\t#{rewriter.map(eval(ast)).inspect}"
      else
        output.puts line
      end
    end
    @count_replaced = rewriter.count_replaced
  end
end
