class ReplaceSharedStringAst 
  
  attr_accessor :shared_strings
  
  def initialize(shared_strings)
    @shared_strings = shared_strings
  end
  
  def map(ast)
    return ast unless ast.is_a?(Array)
    shared_string(ast) if ast.first == :shared_string
    ast.each { |a| map(a) }
    ast
  end
  
  # Format [:shared_string, string_id]
  def shared_string(ast)
    ast.replace([:string,shared_strings[ast[1].to_i]])
  end
end
  

class ReplaceSharedStrings
  
  def self.replace(values,shared_strings,output)
    self.new.replace(values,shared_strings,output)
  end
  
  # Rewrites ast with shared strings to strings
  def replace(values,shared_strings,output)
    rewriter = ReplaceSharedStringAst.new(shared_strings)
    values.each_line do |line|
      # Looks to match shared string lines
      if line =~ /\[:shared_string/
        ref, ast = line.split("\t")
        output.puts "#{ref}\t#{rewriter.map(eval(ast)).inspect}"
      else
        output.puts line
      end
    end
  end
end
