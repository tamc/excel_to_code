require_relative "map_values_to_ruby"

class CompileToRubyUnitTest
  

  attr_accessor :epsilon
  attr_accessor :delta

  def initialize
    @epsilon = 0.002
    @delta = 0.002
  end

  def self.rewrite(*args)
    self.new.rewrite(*args)
  end
  
  def rewrite(input, sloppy, sheet_names,  o)
    mapper = MapValuesToRuby.new
    input.each do |ref, ast|
      c_name = sheet_names[ref.first.to_s] || ref.first.to_s #FIXME: Need to make it the actual c_name
      cell = ref.last
      value = mapper.map(ast)
      full_reference = "worksheet.#{c_name}_#{cell.downcase}"
      test_name = "test_#{c_name}_#{cell.downcase}"
      case ast.first
      when :blank
        if sloppy
          o.puts "  def #{test_name}; assert_includes([nil, 0], #{full_reference}); end"
        else
          o.puts "  def #{test_name}; assert_equal(#{value}, #{full_reference}); end"
        end
      when :number
        if sloppy
          if value.to_f.abs <= 1
            if value == "0" 
              o.puts "  def #{test_name}; assert_in_delta(#{value}, (#{full_reference}||0), #{delta}); end"
            else
              o.puts "  def #{test_name}; assert_in_delta(#{value}, #{full_reference}, #{delta}); end"
            end
          else
            o.puts "  def #{test_name}; assert_in_epsilon(#{value}, #{full_reference}, #{epsilon}); end"
          end
        else
          o.puts "  def #{test_name}; assert_equal(#{value}, #{full_reference}); end"
        end
      else
        o.puts "  def #{test_name}; assert_equal(#{value}, #{full_reference}); end"
      end
    end
  end
  
end
