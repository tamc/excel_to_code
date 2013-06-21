require_relative 'map_formulae_to_c'

class CompileToC
  
  attr_accessor :settable
  attr_accessor :gettable
  attr_accessor :worksheet
  attr_accessor :variable_set_counter
  
  def self.rewrite(*args)
    self.new.rewrite(*args)
  end
  
  def rewrite(input,sheet_names_file,output)
    self.settable ||= lambda { |ref| false }
    self.gettable ||= lambda { |ref| true }
    @variable_set_counter ||= 0
    mapper = MapFormulaeToC.new
    mapper.worksheet = worksheet
    mapper.sheet_names = Hash[sheet_names_file.readlines.map { |line| line.strip.split("\t")}]
    c_name = mapper.sheet_names[worksheet] || worksheet
    input.each_line do |line|
      begin
        ref, formula = line.split("\t")
        ast = eval(formula)
        calculation = mapper.map(ast)
        name = c_name ? "#{c_name}_#{ref.downcase}" : ref.downcase
        static_or_not = gettable.call(ref) ? "" : "static "
        if settable.call(ref)
          output.puts "ExcelValue #{name}_default() {"
          mapper.initializers.each do |i|
            output.puts "  #{i}"
          end
          output.puts "  return #{calculation};"
          output.puts "}"
          output.puts "static ExcelValue #{name}_variable;"
          output.puts "ExcelValue #{name}() { if(variable_set[#{@variable_set_counter}] == 1) { return #{name}_variable; } else { return #{c_name}_#{ref.downcase}_default(); } }"
          output.puts "void set_#{name}(ExcelValue newValue) { variable_set[#{@variable_set_counter}] = 1; #{name}_variable = newValue; }"
        else
          output.puts "#{static_or_not}ExcelValue #{name}() {"
          output.puts "  static ExcelValue result;"
          output.puts "  if(variable_set[#{@variable_set_counter}] == 1) { return result;}"
          mapper.initializers.each do |i|
            output.puts "  #{i}"
          end
          output.puts "  result = #{calculation};"
          output.puts "  variable_set[#{@variable_set_counter}] = 1;"
          output.puts "  return result;"
          output.puts "}"
        end
        output.puts
        @variable_set_counter += 1
        mapper.reset
      rescue Exception => e
        puts "Exception at line #{line}"
        raise
      end      
    end
  end
  
end
