require_relative 'map_formulae_to_ruby'

class CompileToRuby
  
  def self.rewrite(*args)
    self.new.rewrite(*args)
  end
  
  def rewrite(input,output)
    mapper = MapFormulaeToRuby.new
    input.lines do |line|
      ref, formula = line.split("\t")
      output.puts "  def #{ref.downcase}; #{mapper.map(eval(formula))}; end"
    end
  end
  
end