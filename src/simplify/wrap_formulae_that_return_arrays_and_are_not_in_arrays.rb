require_relative '../excel'

class WrapFormulaeThatReturnArraysAndAReNotInArrays
  
  def self.replace(*args)
    self.new.replace(*args)
  end
  
  def replace(input,output)
  
    input.each_line do |line|
      # Looks to match lines that contain formulae that return ranges, such as MMULT
      if line =~ /"MMULT"/
        content = line.split("\t")
        ast = eval(content.pop)
        if ast[0] == :function && ast[1] == "MMULT"
          new_ast = [:function, "INDEX", ast, [:number, "1"], [:number, "1"]] 
          output.puts "#{content.join("\t")}\t#{new_ast.inspect}"          
        else
          output.puts line
        end
      else
        output.puts line
      end
    end
  end
  
end
