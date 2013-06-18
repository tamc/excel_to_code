require_relative "ast_expand_array_formulae"

class RewriteArrayFormulaeToArrays
  
  def self.rewrite(*args)
    new.rewrite(*args)
  end
  
  def rewrite(input,output)
    mapper = AstExpandArrayFormulae.new
    input.each_line do |line|
      content = line.split("\t")
      ast = eval(content.pop)
      output.puts "#{content.join("\t")}\t#{mapper.map(ast).inspect}"
    end
  end
  
end
