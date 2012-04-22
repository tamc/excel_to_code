# coding: utf-8

require_relative 'excel_to_x'

class ExcelToRuby < ExcelToX
  
  def language
    "ruby"
  end  
  
  # Skip this
  def replace_values_with_constants
    
    worksheets("Skipping replacing values with constants") do |name,xml_filename|    
      i = File.join(intermediate_directory, name, "formulae_inlined_pruned_replaced-1.ast")
      o = File.join(intermediate_directory, name, "formulae_inlined_pruned_replaced.ast")
      if run_in_memory
        @files[o] = @files[i]
      else
        `cp '#{i}' '#{o}'`
      end
    end

    i = File.join(intermediate_directory,"common-elements-1.ast")
    o = File.join(intermediate_directory,"common-elements.ast")
    if run_in_memory
      @files[o] = @files[i]
    else
      `cp '#{i}' '#{o}'`
    end
  end
  
  # These actually create the code version of the excel
  def write_code
    write_out_excel_as_code
    write_out_test_as_code
  end
    
  def write_out_excel_as_code
    w = input("worksheet_c_names")
    o = output("#{output_name.downcase}.rb")
    o.puts "# coding: utf-8"
    o.puts "# Compiled version of #{excel_file}"
    o.puts "require '#{File.expand_path(File.join(File.dirname(__FILE__),'../excel/excel_functions'))}'"
    o.puts ""
    o.puts "class #{ruby_module_name}"
    o.puts "  include ExcelFunctions"
    
    o.puts  
    o.puts "  # Starting common elements"
    c = CompileToRuby.new
    i = input("common-elements.ast")
    w.rewind    
    c.rewrite(i,w,o)
    o.puts "  # Ending common elements"
    o.puts
    close(i)
    
    d = intermediate('defaults')
    
    worksheets("Turning worksheet into code") do |name,xml_filename|
      c.settable = settable(name)
      c.worksheet = name
      i = input(name,"formulae_inlined_pruned_replaced.ast")
      w.rewind
      o.puts "  # Start of #{name}"
      c.rewrite(i,w,o,d)
      o.puts "  # End of #{name}"
      o.puts ""
      close(i)
    end   
     
    close(d)
    
    o.puts
    o.puts "  # starting initializer"
    o.puts "  def initialize"
    d = input('defaults')
    d.lines do |line|
      o.puts line
    end
    o.puts "  end"
    o.puts ""
    close(d)
              
    o.puts "end"
    close(w,o)
  end

  def write_out_test_as_code
    o = output("test_#{output_name.downcase}.rb")
    
    o.puts "# coding: utf-8"
    o.puts "# All tests for #{excel_file}"
    o.puts "require 'test/unit'"
    o.puts "require_relative '#{output_name.downcase}'"
    o.puts
    o.puts "class Test#{output_name.capitalize} < Test::Unit::TestCase"
    o.puts "  def worksheet; @worksheet ||= #{ruby_module_name}.new; end"
    
    c = CompileToRubyUnitTest.new
    all_formulae = all_formulae('formulae_inlined_pruned_replaced.ast')
    
    worksheets("Compiling worksheet") do |name,xml_filename|
      i = input(name,"values_pruned2.ast")
      o.puts "  # Start of #{name}"
      c_name = c_name_for_worksheet_name(name)
      if !cells_to_keep || cells_to_keep.empty? || cells_to_keep[name] == :all
        refs_to_test = all_formulae[name].keys
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