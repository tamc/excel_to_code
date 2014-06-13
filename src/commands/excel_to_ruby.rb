# coding: utf-8

require_relative 'excel_to_x'

class ExcelToRuby < ExcelToX
  
  def language
    "ruby"
  end  
  
  # Skip this
  def replace_values_with_constants    
  end
  
  # These actually create the code version of the excel
  def write_code
    write_out_excel_as_code
    write_out_test_as_code
  end
    
  def write_out_excel_as_code
    log.info "Starting to write out code"
    
    o = output("#{output_name.downcase}.rb")

    o.puts "# coding: utf-8"
    o.puts "# Compiled version of #{excel_file}"
    # FIXME: Should include the ruby files as part of the output, so don't have any dependencies
    o.puts "require '#{File.expand_path(File.join(File.dirname(__FILE__),'../excel/excel_functions'))}'"
    o.puts ""
    o.puts "class #{ruby_module_name}"
    o.puts "  include ExcelFunctions"
    o.puts "  def original_excel_filename; #{excel_file.inspect}; end"
    
    c = CompileToRuby.new
    c.settable = settable

    c.rewrite(@formulae, @worksheet_c_names, o)
    o.puts
    
    # Output the named references

    # Getters
    o.puts "# Start of named references"
    c.settable = lambda { |ref| false }
    named_references_ast = {}
    @named_references_to_keep.each do |ref|
      c_name = ref.is_a?(Array) ? c_name_for(ref) : ["", c_name_for(ref)]
      named_references_ast[c_name] = @named_references[ref] || @table_areas[ref]
    end

    c.rewrite(named_references_ast, @worksheet_c_names, o)

    # Setters
    m = MapNamedReferenceToRubySetter.new
    m.cells_that_can_be_set_at_runtime = cells_that_can_be_set_at_runtime
    m.sheet_names = @worksheet_c_names
    @named_references_that_can_be_set_at_runtime.each do |ref|
      c_name = c_name_for(ref)
      ast = @named_references[ref] || @table_areas[ref]
      o.puts "  def #{c_name}=(newValue)"
      o.puts "    @#{c_name} = newValue"
      o.puts m.map(ast)
      o.puts "  end"
    end
    o.puts "# End of named references"

    log.info "Starting to write initializer"
    o.puts
    o.puts "  # starting initializer"
    o.puts "  def initialize"
    d = c.defaults
    d.each do |line|
      o.puts line
    end
    o.puts "  end"
    o.puts ""
    log.info "Finished writing initializer"
              

    o.puts "end"
    close(o)
    log.info "Finished writing code"
  end

  def write_out_test_as_code
    o = output("test_#{output_name.downcase}.rb")
    
    o.puts "# coding: utf-8"
    o.puts "# All tests for #{excel_file}"
    o.puts "require 'minitest/autorun'"
    o.puts "require_relative '#{output_name.downcase}'"
    o.puts
    o.puts "class Test#{ruby_module_name} < Minitest::Unit::TestCase"
    o.puts "  def worksheet; @worksheet ||= #{ruby_module_name}.new; end"

    CompileToCUnitTest.rewrite(Hash[@references_to_test_array], sloppy_tests, @worksheet_c_names, @constants, o)

    o.puts "end"   
    close(o)
  end
  
  def compile_code
    # Not needed
  end
  
  def run_tests
    return unless actually_run_tests
    puts "Running the resulting tests"
    puts `cd #{File.join(output_directory)}; ruby "test_#{output_name.downcase}.rb"`
  end
  
end
