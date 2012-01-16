require_relative 'map_formulae_to_ruby'

class CompileToRuby
  
  attr_accessor :settable
  
  def self.rewrite(*args)
    self.new.rewrite(*args)
  end
  
  def rewrite(input,sheet_names_file,output,defaults = nil)
    self.settable ||= lambda { |ref| false }
    mapper = MapFormulaeToRuby.new
    mapper.sheet_names = Hash[sheet_names_file.readlines.map { |line| line.strip.split("\t")}]
    input.lines do |line|
      ref, formula = line.split("\t")
      if settable.call(ref)
        output.puts "  attr_accessor :#{ref.downcase} # Default: #{mapper.map(eval(formula))}"
        defaults.puts "    @#{ref.downcase} = #{mapper.map(eval(formula))}" if defaults
      else
        output.puts "  def #{ref.downcase}; #{mapper.map(eval(formula))}; end"
      end
    end
  end
  
end