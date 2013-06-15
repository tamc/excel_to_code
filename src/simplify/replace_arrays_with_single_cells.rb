require_relative '../excel'

class ReplaceArraysWithSingleCells
  
  def self.replace(*args)
    self.new.replace(*args)
  end
  
  def replace(input,output)
  
    input.each_line do |line|
      # Looks to match shared string lines
      if line =~ /\[:array/
        content = line.split("\t")
        ast = eval(content.pop)
        if ast.first == :array
          output.puts "#{content.join("\t")}\t#{ast[1][1].inspect}"          
        else
          output.puts line
        end
      else
        output.puts line
      end
    end
  end
  
end
