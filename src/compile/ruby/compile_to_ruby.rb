require_relative 'map_formulae_to_ruby'

class CompileToRuby
  
  attr_accessor :settable
  attr_accessor :worksheet
  
  def self.rewrite(*args)
    self.new.rewrite(*args)
  end
  
  def rewrite(input,sheet_names_file,output,defaults = nil)
    self.settable ||= lambda { |ref| false }
    mapper = MapFormulaeToRuby.new
    mapper.worksheet = worksheet
    mapper.sheet_names = Hash[sheet_names_file.readlines.map { |line| line.strip.split("\t")}]
    c_name = mapper.sheet_names[worksheet]
    input.each_line do |line|
      begin
        ref, formula = line.split("\t")
        name = c_name ? "#{c_name}_#{ref.downcase}" : ref.downcase
        if settable.call(ref)
          output.puts "  attr_accessor :#{name} # Default: #{mapper.map(eval(formula))}"
          defaults.puts "    @#{name} = #{mapper.map(eval(formula))}" if defaults
        else
          output.puts "  def #{name}; @#{name} ||= #{mapper.map(eval(formula))}; end"
        end
      rescue Exception => e
        puts "Exception at line #{line}"
        raise
      end      
    end
  end
  
end
