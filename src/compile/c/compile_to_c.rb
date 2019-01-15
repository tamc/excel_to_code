require_relative 'map_formulae_to_c'
require 'set'

class CompileToC
  
  attr_accessor :settable
  attr_accessor :gettable
  attr_accessor :variable_set_counter
  attr_accessor :variable_set_sheet_hash
  attr_accessor :recursion_prevention_sheet_hash
  attr_accessor :allow_unknown_functions
  
  def self.rewrite(*args)
    self.new.rewrite(*args)
  end

  def init_sheet_hash(sheet_names)
    Hash[sheet_names.map {|sheet_name| [sheet_name, Set.new]}]
  end

  def rewrite(formulae, sheet_names, output)
    self.settable ||= lambda { |ref| false }
    self.gettable ||= lambda { |ref| true }
    @variable_set_counter ||= 0
    @recursion_prevention_counter ||= 0
    @variable_set_sheet_hash ||= init_sheet_hash(sheet_names.values.uniq)
    @recursion_prevention_sheet_hash ||= init_sheet_hash(sheet_names.values.uniq)

    mapper = MapFormulaeToC.new
    mapper.allow_unknown_functions = self.allow_unknown_functions
    mapper.sheet_names = sheet_names
    formulae.each do |ref, ast|
      begin
        worksheet = ref.first
        cell = ref.last
        mapper.worksheet = worksheet
        worksheet_c_name = mapper.sheet_names[worksheet.to_s] || worksheet.to_s
        calculation = mapper.map(ast)
        name = worksheet_c_name.length > 0 ? "#{worksheet_c_name}_#{cell.downcase}" : cell.downcase
        
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
          output.puts "ExcelValue #{name}() { if(variable_set[#{@variable_set_counter}] == 1) { return #{name}_variable; } else { return #{name}_default(); } }"
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
            output.puts "  if(recursion_prevention[#{@recursion_prevention_counter}] == 1) { return result;}"
            output.puts "  recursion_prevention[#{@recursion_prevention_counter}] = 1;"

            mapper.initializers.each do |i|
              output.puts "  #{i}".gsub("#{worksheet_c_name}_#{cell.downcase}()", "ZERO")
            end

            output.puts "  result = #{calculation};"
            output.puts "  variable_set[#{@variable_set_counter}] = 1;"
            output.puts "  recursion_prevention[#{@recursion_prevention_counter}] = 0;"
            output.puts "  return result;"
            output.puts "}"
            output.puts
          end
        end
        @variable_set_sheet_hash[worksheet.to_s.downcase].add(@variable_set_counter)
        @recursion_prevention_sheet_hash[worksheet.to_s.downcase].add(@variable_set_counter)
        @variable_set_counter += 1
        @recursion_prevention_counter += 1
        mapper.reset
      rescue Exception => e
        puts "Exception at #{ref} #{ast}"
        if ref.first == "" # Then it is a common method, helpful to indicate where it comes from
          s = /#{ref.last}/io
          formulae.each do |r, a|
            puts "Referenced in #{r}" if a.to_s =~ s
          end
        end
        raise
      end
    end
  end

  def reset_sheets(sheet_names, output)
    sheet_names = sheet_names.values.uniq
    sheet_names.each do |sheet_name|
      output.puts "void reset_#{sheet_name}()\n{"

      @variable_set_sheet_hash[sheet_name].each do |variable|
        output.puts "  variable_set[#{variable}] = 0;"
      end

      @recursion_prevention_sheet_hash[sheet_name].each do |variable|
        output.puts "  recursion_prevention[#{variable}] = 0;"
      end

      output.puts "}\n"
    end
  end
end
