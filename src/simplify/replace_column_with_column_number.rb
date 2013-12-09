class ReplaceColumnWithColumnNumberAST
    
  attr_accessor :count_replaced
  attr_accessor :replacement_made

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
  
  REF_TYPES = {:cell => true, :sheet_reference => true}

  def function(ast)
    return unless ast[1] == :COLUMN
    return unless ast.size == 3
    return unless REF_TYPES.has_key?(ast[2][0])
    if ast[2][0] == :cell
      reference = Reference.for(ast[2][1])
    elsif ast[2][0] == :sheet_reference
      reference = Reference.for(ast[2][2][1])
    end
    reference.calculate_excel_variables
    @count_replaced += 1
    @replacement_made = true
    ast.replace( CachingFormulaParser.map([:number, reference.excel_column_number]))
  end

end
  

class ReplaceColumnWithColumnNumber
    
  def self.replace(*args)
    self.new.replace(*args)
  end

  attr_accessor :count_replaced
  
  def replace(input,output)
    rewriter = ReplaceColumnWithColumnNumberAST.new
    input.each_line do |line|
      # Looks to match lines with references
      if line =~ /:COLUMN/
        ref, ast = line.split("\t")
        output.puts "#{ref}\t#{rewriter.map(eval(ast)).inspect}"
      else
        output.puts line
      end
    end
    @count_replaced = rewriter.count_replaced
  end
end
