class ReplaceOffsetsWithReferencesAst

  attr_accessor :replacements_made_in_the_last_pass

  def initialize
    @replacements_made_in_the_last_pass = 0
  end
    
  def map(ast)
    return ast unless ast.is_a?(Array)
    operator = ast[0]
    if respond_to?(operator)
      send(operator,*ast[1..-1])
    else
      [operator,*ast[1..-1].map {|a| map(a) }]
    end
  end
  
  def function(name,*args)
    if name == "OFFSET"
      try_to_replace_offset(*args)
    else
      [:function,name,*args.map { |a| map(a) }]
    end
  end

  def try_to_replace_offset(reference, row_offset, column_offset, height = [:number, 1], width = [:number, 1])
    if [row_offset, column_offset, height, width].all? { |a| a.first == :number }
      if reference.first == :cell
        offset_cell(reference, row_offset, column_offset, height, width)
      elsif reference.first == :sheet_reference && reference[2].first == :cell
        [:sheet_reference, reference[1], offset_cell(reference[2], row_offset, column_offset, height, width)]
      else
        puts "#{[:function, "OFFSET", reference, row_offset, column_offset, height, width]} not replaced"
        [:function, "OFFSET", reference, row_offset, column_offset, height, width]
      end
    else
        puts "#{[:function, "OFFSET", reference, row_offset, column_offset, height, width]} not replaced"
        [:function, "OFFSET", reference, row_offset, column_offset, height, width]
    end
  end

  def offset_cell(reference, row_offset, column_offset, height, width)

    reference = reference[1]
    row_offset = row_offset[1].to_i
    column_offset = column_offset[1].to_i
    height = height[1].to_i
    width = width[1].to_i

    @replacements_made_in_the_last_pass += 1
    reference = Reference.for(reference.gsub("$",""))
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
  
  attr_accessor :replacements_made_in_the_last_pass
  
  def replace(input,output)
    rewriter = ReplaceOffsetsWithReferencesAst.new
    input.each_line do |line|
      # Looks to match lines with references
      if line =~ /"OFFSET"/
        ref, ast = line.split("\t")
        output.puts "#{ref}\t#{rewriter.map(eval(ast)).inspect}"
      else
        output.puts line
      end
    end
    @replacements_made_in_the_last_pass = rewriter.replacements_made_in_the_last_pass
  end
end
