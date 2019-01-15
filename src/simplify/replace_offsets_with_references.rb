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
    height = ast[5]
    width = ast[6]

    reference = reference.original unless %i{cell sheet_reference array}.include?(reference.first)
    
    unless height
      if reference.first != :array
        height = [:number, 1.0]
      else
        height = [:number, reference.length - 1]
      end
    end

    unless width
      if reference.first != :array
        width = [:number, 1.0]
      else
        width = [:number, reference[1].length - 1]
      end
    end

    if reference.first == :array
      reference = reference[1][1].original
    end

    [row_offset, column_offset, height, width].each do |arg|
       next unless arg.first == :error
       ast.replace(arg) 
       return
    end

    return unless [row_offset, column_offset, height, width].all? { |a| a.first == :number }

    if reference.first == :cell
      ast.replace(offset_cell(reference, row_offset, column_offset, height, width, nil))
    elsif reference.first == :sheet_reference && reference[2].first == :cell
      ast.replace(offset_cell(reference[2], row_offset, column_offset, height, width, reference[1]))
    else
      p "OFFSET reference is #{reference} from #{ast}, so not replacing"
    end
  end

  def offset_cell(reference, row_offset, column_offset, height, width, sheet)

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
      if sheet 
        return [:sheet_reference, sheet, [:cell, start_reference]]
      else
        return [:cell, start_reference]
      end
    else
      area = Area.for("#{start_reference}:#{end_reference}")
      return area.to_array_literal(sheet)
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
