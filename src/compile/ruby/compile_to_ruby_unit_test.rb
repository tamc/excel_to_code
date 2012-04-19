require_relative "map_values_to_ruby"

class CompileToRubyUnitTest
  
  def self.rewrite(*args)
    self.new.rewrite(*args)
  end
  
  def rewrite(input, c_name, refs_to_test, output)
    mapper = MapValuesToRuby.new
    input.lines do |line|
      ref, formula = line.split("\t")
      next unless refs_to_test.include?(ref.downcase)
      ast = eval(formula)
      value = mapper.map(ast)
      full_reference = "worksheet.#{c_name}_#{ref.downcase}"
      if ast.first == :number
        if value == "0" # Need to do a slightly different test, because needs to pass if nil returned, as well as zero
          output.puts "  def test_#{c_name}_#{ref.downcase}; assert_in_epsilon(#{value},#{full_reference} || 0); end"      
        else
          output.puts "  def test_#{c_name}_#{ref.downcase}; assert_in_epsilon(#{value},#{full_reference}); end"      
        end
      else
        output.puts "  def test_#{c_name}_#{ref.downcase}; assert_equal(#{value},#{full_reference}); end"
      end
    end
  end
  
end