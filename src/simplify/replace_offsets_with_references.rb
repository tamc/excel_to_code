class ReplaceOffsetsWithReferencesAst

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
  
  def function(ast)
    name = ast[1]
    args = ast[2..-1]
    return unless ast[1] == :OFFSET
    reference = ast[2]
    row_offset = ast[3]
    column_offset = ast[4]
    height = ast[5] || [:number, 1]
    width = ast[6] || [:number, 1]
    return unless [row_offset, column_offset, height, width].all? { |a| a.first == :number }
    if reference.first == :cell
      ast.replace(offset_cell(reference, row_offset, column_offset, height, width))
    elsif reference.first == :sheet_reference && reference[2].first == :cell
      ast.replace([:sheet_reference, reference[1], offset_cell(reference[2], row_offset, column_offset, height, width)])
    end
  end

  def offset_cell(reference, row_offset, column_offset, height, width)

    reference = reference[1]
    row_offset = row_offset[1].to_i
    column_offset = column_offset[1].to_i
    height = height[1].to_i
    width = width[1].to_i

    @count_replaced += 1
    @replacement_made = true

    reference = Reference.for(reference).unfix
    start_reference = reference.offset(row_offset.to_i, column_offset.to_i)
    end_reference = reference.offset(row_offset.to_i + height.to_i - 1, column_offset.to_i + width.to_i - 1)
    if start_reference == end_reference
      [:cell, start_reference]
    else
      [:area, start_reference, end_reference]
    end
  end

end
  

class ReplaceOffsetsWithReferences
    
  def self.replace(*args)
    self.new.replace(*args)
  end
  
  attr_accessor :count_replaced
  
  def replace(input,output)
    rewriter = ReplaceOffsetsWithReferencesAst.new
    input.each_line do |line|
      # Looks to match lines with references
      if line =~ /:OFFSET/
        ref, ast = line.split("\t")
        output.puts "#{ref}\t#{rewriter.map(eval(ast)).inspect}"
      else
        output.puts line
      end
    end
    @count_replaced = rewriter.count_replaced
  end
end
