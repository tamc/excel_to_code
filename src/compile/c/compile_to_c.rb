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
          output.puts ast_to_c_setter(ast,name)
          output.puts "ExcelValue #{name}() { return #{name}_variable; }"
          defaults.puts "  #{name}_variable = #{calculation};" if defaults
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
  
  def ast_to_c_setter(ast,name)
    case ast.first
    when :number, :percentage
      "void set_#{name}(double number) { #{name}_variable = new_excel_number(number); }"
    when :error
      "void set_#{name}(int error_number) { }"
    when :string
      "void set_#{name}(char *string) { #{name}_variable = new_excel_string(string); }"      
    when :boolean_true, :boolean_false
      "void set_#{name}(int true_or_false) { #{name}_variable = (true_or_false == 1) ?  TRUE : FALSE; }"
    else
      raise NotSupportedException.new("#{ast} type can't be settable")
    end
  end
  
end