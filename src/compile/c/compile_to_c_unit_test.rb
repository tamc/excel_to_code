class CompileToCUnitTest

  attr_accessor :epsilon
  attr_accessor :delta

  def initialize
    @epsilon = 0.001
    @delta = 0.001
  end
  
  def self.rewrite(*args)
    self.new.rewrite(*args)
  end
  
  def rewrite(input, sloppy, c_name, refs_to_test, output)
    input.lines do |line|
      begin
        ref, formula = line.split("\t")
        next unless refs_to_test.include?(ref.upcase)
        output.puts "def test_#{c_name}_#{ref.downcase}"
        output.puts "  r = spreadsheet.#{c_name}_#{ref.downcase}"
        ast = eval(formula)
        case ast.first
        when :number, :percentage
          unless sloppy
            output.puts "  assert_equal(:ExcelNumber,r[:type])"
            output.puts "  assert_equal(#{ast.last.to_f.to_s},r[:number])"
          else
            if ast.last.to_f == 0
              output.puts "  pass if r[:type] == :ExcelEmpty"
            end
            
            output.puts "  assert_equal(:ExcelNumber,r[:type])"

            if ast.last.to_f <= 1
              output.puts "  assert_in_delta(#{ast.last.to_f.to_s},r[:number],#{@delta})"
            else
              output.puts "  assert_in_epsilon(#{ast.last.to_f.to_s},r[:number],#{@epsilon})"
            end
          end
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
        when :blank
          unless sloppy
            output.puts "  assert_equal(:ExcelEmpty,r[:type])"
          else
            output.puts "  pass if r[:type] == :ExcelEmpty"
            output.puts "  assert_equal(:ExcelNumber,r[:type])"
            output.puts "  assert_in_delta(0.0,r[:number],#{@delta})"
          end
        else
          raise NotSupportedException.new("#{ast} type can't be tested")
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
