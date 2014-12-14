require_relative "map_values_to_c"

class CompileToCUnitTest

  def self.rewrite(*args)
    self.new.rewrite(*args)
  end
  
  def rewrite(input, sloppy, sheet_names, constants,  o)
    mapper = MapValuesToC.new
    input.each do |ref, ast|
      worksheet_c_name = sheet_names[ref.first.to_s] || ref.first.to_s #FIXME: Need to make it the actual c_name
      cell = ref.last
      value = mapper.map(ast)
      full_reference = worksheet_c_name.length > 0 ? "#{worksheet_c_name}_#{cell.downcase}()" : "#{cell.downcase}()"
      test_name = "'#{ref.first.to_s}'!#{cell.upcase}"
      o.puts "  assert_equal(#{value}, #{full_reference}, #{test_name.inspect});"
    end
  end
  
end
