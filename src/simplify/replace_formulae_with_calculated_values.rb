require_relative 'map_formulae_to_values'

class ReplaceFormulaeWithCalculatedValues
    
  def self.replace(*args)
    self.new.replace(*args)
  end
  
  def replace(input,output)
    rewriter = MapFormulaeToValues.new
    input.lines do |line|
      ref, ast = line.split("\t")
      output.puts "#{ref}\t#{rewriter.map(eval(ast)).inspect}"
    end
  end
end
