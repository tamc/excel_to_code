require_relative 'map_formulae_to_c'

class CompileToC
  
  attr_accessor :settable
  attr_accessor :gettable
  attr_accessor :variable_set_counter
  
  def self.rewrite(*args)
    self.new.rewrite(*args)
  end
  
  def rewrite(formulae, sheet_names, output)
    self.settable ||= lambda { |ref| false }
    self.gettable ||= lambda { |ref| true }
    @variable_set_counter ||= 0

    mapper = MapFormulaeToC.new
    mapper.sheet_names = sheet_names
    formulae.each do |ref, ast|
      begin
        worksheet = ref.first
        cell = ref.last
        mapper.worksheet = worksheet
        c_name = mapper.sheet_names[worksheet.to_s] || worksheet.to_s
        calculation = mapper.map(ast)
        name = c_name ? "#{c_name}_#{cell.downcase}" : cell.downcase
        
        # Declare function as static so it can be inlined where possible
        static_or_not = gettable.call(ref) ? "" : "static "

        # Settable functions need a default value and then a function for getting and setting
        if settable.call(ref)
          output.puts "ExcelValue #{name}_default() {"
          mapper.initializers.each do |i|
            output.puts "  #{i}"
          end
          output.puts "  return #{calculation};"
          output.puts "}"
          output.puts "static ExcelValue #{name}_variable;"
          output.puts "ExcelValue #{name}() { if(variable_set[#{@variable_set_counter}] == 1) { return #{name}_variable; } else { return #{c_name}_#{cell.downcase}_default(); } }"
          output.puts "void set_#{name}(ExcelValue newValue) { variable_set[#{@variable_set_counter}] = 1; #{name}_variable = newValue; }"
          output.puts
        # Other functions just have a getter
        else
          # In simple cases, don't bother memoizing the result
          simple = (ast[0] == :constant) || (ast[0] == :cell && ast[1] =~ /common\d+/) || (ast[0] == :blank) || (ast[0] == :error)

          if simple
            output.puts "#{static_or_not}ExcelValue #{name}() { return #{calculation}; }"
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
            output.puts
          end
        end
        @variable_set_counter += 1
        mapper.reset
      rescue Exception => e
        puts "Exception at #{ref} #{ast}"
        raise
      end      
    end
  end
  
end
