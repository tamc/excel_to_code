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
    
    w = input('Worksheet C names')
    o = output("#{output_name.downcase}.rb")
    o.puts "# coding: utf-8"
    o.puts "# Compiled version of #{excel_file}"
    # FIXME: Should include the ruby files as part of the output, so don't have any dependencies
    o.puts "require '#{File.expand_path(File.join(File.dirname(__FILE__),'../excel/excel_functions'))}'"
    o.puts ""
    o.puts "class #{ruby_module_name}"
    o.puts "  include ExcelFunctions"
    
    o.puts  
    o.puts "  # Starting common elements"
    log.info "Starting to write code for common elements"
    c = CompileToRuby.new
    i = input("Common elements")
    w.rewind    
    c.rewrite(i,w,o)
    o.puts "  # Ending common elements"
    o.puts
    close(i)
    log.info "Finished writing code for common elements"
    
    d = intermediate('Defaults')
    
    worksheets do |name,xml_filename|
      log.info "Starting to write code for worksheet #{name}"
      c.settable = settable(name)
      c.worksheet = name
      i = input([name,"Formulae"])
      w.rewind
      o.puts "  # Start of #{name}"
      c.rewrite(i,w,o,d)
      o.puts "  # End of #{name}"
      o.puts ""
      close(i)
      log.info "Finished writing code for worksheet #{name}"
    end   
     
    close(d)
    
    log.info "Starting to write initializer"
    o.puts
    o.puts "  # starting initializer"
    o.puts "  def initialize"
    d = input('Defaults')
    d.lines do |line|
      o.puts line
    end
    o.puts "  end"
    o.puts ""
    close(d)
    log.info "Finished writing initializer"
              
    o.puts "end"
    close(w,o)
    log.info "Finished writing code"
  end

  def write_out_test_as_code
    o = output("test_#{output_name.downcase}.rb")
    
    o.puts "# coding: utf-8"
    o.puts "# All tests for #{excel_file}"
    o.puts "require 'test/unit'"
    o.puts "require_relative '#{output_name.downcase}'"
    o.puts
    o.puts "class Test#{ruby_module_name} < Test::Unit::TestCase"
    o.puts "  def worksheet; @worksheet ||= #{ruby_module_name}.new; end"
    
    c = CompileToRubyUnitTest.new
    formulae = all_formulae()
    
    worksheets do |name,xml_filename|
      i = input(name,"Values")
      o.puts "  # Start of #{name}"
      c_name = c_name_for_worksheet_name(name)
      if !cells_to_keep || cells_to_keep.empty? || cells_to_keep[name] == :all
        refs_to_test = formulae[name].keys
      else
        refs_to_test = cells_to_keep[name]
      end
      if refs_to_test && !refs_to_test.empty?
        refs_to_test = refs_to_test.map(&:upcase)
        c.rewrite(i, c_name, refs_to_test, o)
      end
      o.puts "  # End of #{name}"
      o.puts ""
      close(i)
    end 
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