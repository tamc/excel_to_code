require_relative 'map_formulae_to_c_function'

class CompileToCFunction
  
  attr_accessor :settable
  attr_accessor :gettable
  attr_accessor :variable_set_counter
  
  def self.rewrite(*args)
    self.new.rewrite(*args)
  end
  
  def rewrite(formulae, sheet_names, output, order = formulae.keys)
    @variable_set_counter ||= 0

    mapper = MapFormulaeToCFunction.new
    mapper.sheet_names = sheet_names

    order.each do |ref|
      ast = formulae[ref]
      next unless ast
      begin
        worksheet = ref.first
        cell = ref.last
        mapper.worksheet = worksheet
        mapper.counter = @variable_set_counter

        worksheet_c_name = mapper.sheet_names[worksheet.to_s] || worksheet.to_s
        calculation = mapper.map(ast)
        name = worksheet_c_name.length > 0 ? "#{worksheet_c_name}_#{cell.downcase}" : cell.downcase

        mapper.initializers.each do |i|
          output.puts "  #{i}".gsub(/^  static ExcelValue/,'  ExcelValue')
        end

        output.puts "  ExcelValue #{name} = #{calculation};"
        @variable_set_counter = mapper.counter
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
  
end
