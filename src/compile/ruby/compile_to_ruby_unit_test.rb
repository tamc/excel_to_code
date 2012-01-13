require_relative "map_values_to_ruby"

class CompileToRubyUnitTest
  
  def self.rewrite(*args)
    self.new.rewrite(*args)
  end
  
  def rewrite(input,output)
    mapper = MapValuesToRuby.new
    input.lines do |line|
      ref, formula = line.split("\t")
      output.puts "  def test_#{ref.downcase}; assert_equal(worksheet.#{ref.downcase},#{mapper.map(eval(formula))}); end"
    end
  end
  
end