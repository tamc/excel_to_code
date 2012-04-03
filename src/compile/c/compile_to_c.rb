require_relative 'map_formulae_to_c'

class CompileToC
  
  attr_accessor :settable
  attr_accessor :worksheet
  
  def self.rewrite(*args)
    self.new.rewrite(*args)
  end
  
  def rewrite(input,sheet_names_file,output)
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
        name = "#{c_name}_#{ref.downcase}"
        if settable.call(ref)
          output.puts "ExcelValue #{name}_default() {"
          output.puts "  static ExcelValue result;"
          output.puts "  static int calculated = 0;"
          output.puts "  if(calculated == 1) { return result;}"
            mapper.initializers.each do |i|
              output.puts "  #{i}"
            end
            #output.puts "  return #{calculation};"
          output.puts "  result = #{calculation};"
          output.puts "  calculated = 1;"
          output.puts "  return result;"
          output.puts "}"
          output.puts "static ExcelValue #{name}_variable;"
          output.puts "static int #{name}_variable_set = 0;"
          output.puts "ExcelValue #{name}() { if(#{name}_variable_set == 1) { return #{name}_variable; } else { return #{c_name}_#{ref.downcase}_default(); } }"
          output.puts "void set_#{name}(ExcelValue newValue) { #{name}_variable_set = 1; #{name}_variable = newValue; }"
        else
          output.puts "ExcelValue #{name}() {"
          output.puts "  static ExcelValue result;"
          output.puts "  static int calculated = 0;"
          output.puts "  if(calculated == 1) { return result;}"
            mapper.initializers.each do |i|
              output.puts "  #{i}"
            end
            #output.puts "  return #{calculation};"
          output.puts "  result = #{calculation};"
          output.puts "  calculated = 1;"
          output.puts "  return result;"
          output.puts "}"
        end
        mapper.reset
      rescue Exception => e
        puts "Exception at line #{line}"
        raise
      end      
    end
  end
  
end