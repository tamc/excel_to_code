require_relative '../excel'

class ReplaceRangesWithArrayLiteralsAst
  def map(ast)
    if ast.is_a?(Array)
      operator = ast.shift
      if respond_to?(operator)
        send(operator,*ast)
      else
        [operator,*ast.map {|a| map(a) }]
      end
    else
      return ast
    end
  end
  
  def sheet_reference(sheet,reference)
    if reference.first == :area
      area = Area.for("#{reference[1]}:#{reference[2]}")
      area.to_array_literal(sheet)
    else
      [:sheet_reference,sheet,reference]
    end
  end
  
  def area(start,finish)
    area = Area.for("#{start}:#{finish}")
    area.to_array_literal
  end
  
end

class ReplaceRangesWithArrayLiterals
  
  def self.replace(*args)
    self.new.replace(*args)
  end
  
  def replace(input,output)
    rewriter = ReplaceRangesWithArrayLiteralsAst.new
  
    input.lines do |line|
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
