require_relative '../excel'

class ReplaceArraysWithSingleCellsAst

  def map(ast)
    return ast unless ast.first == :array
    ast[1][1]
  end
end


class ReplaceArraysWithSingleCells
  
  def self.replace(*args)
    self.new.replace(*args)
  end
  
  def replace(input,output)
  
    replacer = ReplaceArraysWithSingleCellsAst.new
    input.each_line do |line|
      # Looks to match shared string lines
      if line =~ /\[:array/
        content = line.split("\t")
        ast = eval(content.pop)
        output.puts "#{content.join("\t")}\t#{replacer.map(ast).inspect}"          
      else
        output.puts line
      end
    end
  end
  
end
