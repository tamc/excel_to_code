require_relative 'map_formulae_to_c'

class CompileToC
  
  attr_accessor :settable
  attr_accessor :worksheet
  
  def self.rewrite(*args)
    self.new.rewrite(*args)
  end
  
  def rewrite(input,sheet_names_file,output,defaults = nil)
    self.settable ||= lambda { |ref| false }
    mapper = MapFormulaeToC.new
    mapper.worksheet = worksheet
    mapper.sheet_names = Hash[sheet_names_file.readlines.map { |line| line.strip.split("\t")}]
    c_name = mapper.sheet_names[worksheet]
    input.lines do |line|
      begin
        ref, formula = line.split("\t")
        ast = eval(formula)
        calculation = mapper.map(ast)
        if settable.call(ref)
          name = "#{c_name}_#{ref.downcase}"
          output.puts "static ExcelValue #{name}_variable;"
          output.puts "ExcelValue #{name}() { return #{name}_variable; }"
          if mapper.initializers.empty?
            output.puts "ExcelValue #{c_name}_#{ref.downcase}_default() { return #{calculation}; }"
          else
            output.puts "ExcelValue #{c_name}_#{ref.downcase}_default() {"
              mapper.initializers.each do |i|
                output.puts "  #{i}"
              end
              output.puts "  return #{calculation};"
            output.puts "}"
          end
          output.puts "void set_#{name}(ExcelValue newValue) { #{name}_variable = newValue; }"
          defaults.puts "  #{name}_variable = #{c_name}_#{ref.downcase}_default();" if defaults
        else
          if mapper.initializers.empty?
            output.puts "ExcelValue #{c_name}_#{ref.downcase}() { return #{calculation}; }"
          else
            output.puts "ExcelValue #{c_name}_#{ref.downcase}() {"
              mapper.initializers.each do |i|
                output.puts "  #{i}"
              end
              output.puts "  return #{calculation};"
            output.puts "}"
          end
        end
        mapper.reset
      rescue Exception => e
        puts "Exception at line #{line}"
        raise
      end      
    end
  end
  
end