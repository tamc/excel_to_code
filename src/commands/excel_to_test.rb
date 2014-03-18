# coding: utf-8
require_relative 'excel_to_x'
require 'ffi'

class ExcelToTest < ExcelToX
 
  attr_accessor :test_name

  def language
    "C"
  end

  def go!
    # This sorts out the settings
    set_defaults

    log.info "Excel to Code version #{ExcelToCode.version}\n\n"
    
    # These turn the excel into xml on disk
    sort_out_output_directories
    unzip_excel
    
    # These gets the named references, worksheet names and shared strings out of the excel
    extract_data_from_workbook
    
    # This checks that the user inputs of which cells to keep are in the right
    # format and refer to sheets and references that actually exist
    clean_cells_that_can_be_set_at_runtime
    clean_cells_to_keep
    clean_named_references_to_keep
    clean_named_references_that_can_be_set_at_runtime
    
    # This turns named references that are specified as getters and setters
    # into a series of required cell references
    transfer_named_references_to_keep_into_cells_to_keep
    transfer_named_references_that_can_be_set_at_runtime_into_cells_that_can_be_set_at_runtime

    # This makes sure we only extract values from the worksheets we actuall care about
    extractor = ExtractDataFromWorksheet.new
    extractor.only_extract_values = true

    cells_to_keep.each do |name, refs|
      xml_filename = @worksheet_xmls[name]

      log.info "Extracting data from #{name}"
      xml(xml_filename) do |input|
        extractor.extract(name, input)
      end
    end
    @values = extractor.values

    
    # These perform some translations to tsimplify the excel
    # Including:
    # * Turning row and column references (e.g., A:A) to areas, based on the size of the worksheet
    # * Turning range references (e.g., A1:B2) into array litterals (e.g., {A1,B1;A2,B2})
    # * Turning shared formulae into a series of conventional formulae
    # * Turning array formulae into a series of conventional formulae
    # * Mergining all the different types of formulae and values into a single hash
    rewrite_values_to_remove_shared_strings
    
    # FIXME: Bodge for the moment
    @formulae = @values 
    create_sorted_references_to_test

    # This actually creates the code (implemented in subclasses)
    write_code
    
    # These compile and run the code version of the excel (implemented in subclasses)
    run_tests
    
    log.info "The generated code is available in #{File.join(output_directory)}"
  end
  
  # These actually create the code version of the excel
  def write_code
    write_tests
  end
    
  def write_code_to_set_values
    log.info "Writing code to set values to match those in the worksheet"
  end
  
  # FIXME: Should make a Rakefile, especially in order to make sure the dynamic library name
  
  def write_tests
    log.info "Writing tests" 

    name = output_name.downcase
    @test_name = "test_#{name}_#{Time.now.to_i}.rb"
    o = output(@test_name)    
    o.puts "# coding: utf-8"
    o.puts "# Test for #{name}"
    o.puts  "require 'minitest/autorun'"
    o.puts  "require_relative '#{output_name.downcase}'"
    o.puts
    o.puts "class Test#{ruby_module_name} < Minitest::Unit::TestCase"
    o.puts "  def self.runnable_methods"
    o.puts "    puts 'Overriding minitest to run tests in a defined order'"
    o.puts "    methods = methods_matching(/^test_/)"
    o.puts "  end" 
    o.puts "  def worksheet; @worksheet ||= init_spreadsheet; end"

    # We need to set up the worksheet with the correct starting values
    o.puts "  def init_spreadsheet"
    o.puts "    s = #{ruby_module_name}Shim.new"

    mapper = MapValuesToRuby.new

    cells_that_can_be_set_at_runtime.each do |sheet, cells|
      worksheet_c_name = c_name_for_worksheet_name(sheet)
      cells.each do |cell|
        ast = @values[[sheet.to_sym, cell.to_sym]]
        next unless ast
        value = mapper.map(ast)

        full_reference = worksheet_c_name.length > 0 ? "#{worksheet_c_name}_#{cell.downcase}" : "#{cell.downcase}"
        o.puts "    s.#{full_reference} = #{value}"
      end
    end

    o.puts "    return s"
    o.puts "  end"
    
    CompileToCUnitTest.rewrite(Hash[@references_to_test_array], sloppy_tests, @worksheet_c_names, @constants, o)
    o.puts "end"
    close(o)
    puts "New test is in #{File.join(output_directory, @test_name)}"
  end

  def run_tests
    return unless actually_run_tests
    puts "Running the resulting tests"
    puts `cd #{File.join(output_directory)}; ruby "#{@test_name}"`
  end
  
end
