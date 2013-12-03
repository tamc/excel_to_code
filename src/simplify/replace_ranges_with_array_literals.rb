require_relative '../excel'

class ReplaceRangesWithArrayLiteralsAst
  def map(ast)
    return ast unless ast.is_a?(Array)
    case ast[0]
    when :sheet_reference; sheet_reference(ast)
    when :area; area(ast)
    end
    ast.each {|a| map(a) }
    ast
  end
  
  # Of the form [:sheet_reference, sheet, reference]
  def sheet_reference(ast)
    sheet = ast[1]
    reference = ast[2]
    return unless reference.first == :area
    area = Area.for("#{reference[1]}:#{reference[2]}")
    a = area.to_array_literal(sheet)
    
    # Don't convert single cell ranges
    if a.size == 2 && a[1].size == 2
      ast.replace(a[1][1])
    else
      ast.replace(a)
    end
  end
  
  # Of the form [:area, start, finish]
  def area(ast)
    start = ast[1]
    finish = ast[2]
    area = Area.for("#{start}:#{finish}")
    a = area.to_array_literal

    # Don't convert single cell ranges
    if a.size == 2 && a[1].size == 2
      ast.replace(a[1][1])
    else
      ast.replace(a)
    end
  end
  
end

class ReplaceRangesWithArrayLiterals
  
  def self.replace(*args)
    self.new.replace(*args)
  end
  
  def replace(input,output)
    rewriter = ReplaceRangesWithArrayLiteralsAst.new
  
    input.each_line do |line|
      # Looks to match shared string lines
      if line =~ /\[:area/
        content = line.split("\t")
        ast = eval(content.pop)
        output.puts "#{content.join("\t")}\t#{rewriter.map(ast).inspect}"
      else
        output.puts line
      end
    end
  end
  
end
