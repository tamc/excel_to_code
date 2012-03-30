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
          output.puts "  assert_equal(r[:type],:ExcelNumber)"
          output.puts "  assert_in_epsilon(r[:number],#{ast.last.to_f.to_s})"
        when :error
        when :string
        when :boolean_true, :boolean_false
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
