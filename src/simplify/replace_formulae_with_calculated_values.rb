require_relative 'map_formulae_to_values'

class ReplaceFormulaeWithCalculatedValues
    
  def self.replace(*args)
    self.new.replace(*args)
  end
  
  def replace(input,output)
    rewriter = MapFormulaeToValues.new
    input.each_line do |line|
      begin
        ref, ast = line.split("\t")
        output.puts "#{ref}\t#{rewriter.map(eval(ast)).inspect}"
        rewriter.reset
      rescue Exception => e
        puts "Exception at line #{line}"
        raise
      end
    end
  end
end
