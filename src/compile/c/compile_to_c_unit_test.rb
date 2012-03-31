class CompileToCUnitTest
  
  def self.rewrite(*args)
    self.new.rewrite(*args)
  end
  
  attr_accessor :worksheet
  
  def rewrite(input,c_name,output,defaults = nil)
    input.lines do |line|
      begin
        ref, formula = line.split("\t")
        output.puts "def test_#{c_name}_#{ref.downcase}"
        output.puts "  r = spreadsheet.#{c_name}_#{ref.downcase}"
        ast = eval(formula)
        case ast.first
        when :number, :percentage
          output.puts "  assert_equal(:ExcelNumber,r[:type])"
          output.puts "  assert_in_epsilon(#{ast.last.to_f.to_s},r[:number])"
        when :error
          output.puts "  assert_equal(:ExcelError,r[:type])"
        when :string
          output.puts "  assert_equal(:ExcelString,r[:type])"
          output.puts "  assert_equal(#{ast.last.inspect},r[:string].force_encoding('utf-8'))" 
        when :boolean_true
          output.puts "  assert_equal(:ExcelBoolean,r[:type])"
          output.puts "  assert_equal(1,r[:number])" 
        when :boolean_false
          output.puts "  assert_equal(:ExcelBoolean,r[:type])"
          output.puts "  assert_equal(0,r[:number])"           
        else
          raise NotSupportedException.new("#{ast} type can't be settable")
        end
        output.puts "end"
        output.puts
      rescue Exception => e
        puts "Exception at line #{line}"
        raise
      end      
    end
  end
  
end
