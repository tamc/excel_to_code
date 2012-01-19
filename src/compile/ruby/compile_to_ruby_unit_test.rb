require_relative "map_values_to_ruby"

class CompileToRubyUnitTest
  
  def self.rewrite(*args)
    self.new.rewrite(*args)
  end
  
  def rewrite(input,output)
    mapper = MapValuesToRuby.new
    input.lines do |line|
      ref, formula = line.split("\t")
      value = mapper.map(eval(formula))
      full_reference = "worksheet.#{ref.downcase}"
      if value == 0 # Need to do a slightly different test, because needs to pass if nil returned, as well as zero
        output.puts "  def test_#{ref.downcase}; assert_equal(#{value},#{full_reference} || 0); end"      
      else
        output.puts "  def test_#{ref.downcase}; assert_equal(#{value},#{full_reference}); end"
      end
    end
  end
  
end