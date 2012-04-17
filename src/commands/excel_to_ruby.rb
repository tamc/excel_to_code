# coding: utf-8

require_relative 'excel_to_x'

class ExcelToRuby < ExcelToX
    
  # These actually create the code version of the excel
  def write_code
    write_out_excel_workbook_as_code
    write_out_excel_workbook_test_as_code
    worksheets("Compiling worksheet") do |name,xml_filename|
      #fork do 
        compile_worksheet_code(name,xml_filename)
        compile_worksheet_test(name,xml_filename)
      #end
    end    
  end
    
  def write_out_excel_workbook_as_code
    w = input("worksheet_ruby_names")
    o = ruby("#{compiled_module_name.downcase}.rb")
    o.puts "# Compiled version of #{excel_file}"
    o.puts "require '#{File.expand_path(File.join(File.dirname(__FILE__),'../excel/excel_functions'))}'"
    o.puts ""
    o.puts "module #{compiled_module_name}"
    o.puts "class Spreadsheet"
    o.puts "  include ExcelFunctions"
    w.lines do |line|
      name, ruby_name = line.strip.split("\t")
      o.puts "  def #{ruby_name}; @#{ruby_name} ||= #{ruby_name.capitalize}.new; end"
    end
    
    c = CompileToRuby.new
    i = input("common-elements.ast")
    w.rewind
    c.rewrite(i,w,o)
          
    o.puts "end"
    o.puts 'Dir[File.join(File.dirname(__FILE__),"worksheets/","*.rb")].each {|f| autoload(File.basename(f,".rb").capitalize,f)}'
    o.puts "end"
    close(i,w,o)
  end

  def write_out_excel_workbook_test_as_code
    w = input("worksheet_ruby_names")
    o = ruby("test_#{compiled_module_name.downcase}.rb")
    o.puts "# All tests for #{excel_file}"
    o.puts  "require 'test/unit'"
    w.lines do |line|
      name, ruby_name = line.strip.split("\t")
      o.puts "require_relative 'tests/test_#{ruby_name.downcase}'"
    end
    close(w,o)
  end
    
  def compile_worksheet_code(name,xml_filename)
    settable_refs = @values_that_can_be_set_at_runtime[name]    
    c = CompileToRuby.new
    c.settable =lambda { |ref| (settable_refs == :all) ? true : settable_refs.include?(ref) } if settable_refs
    c.worksheet = name
    i = input(name,"formulae_inlined_pruned_replaced.ast")
    w = input("worksheet_ruby_names")
    ruby_name = ruby_name_for_worksheet_name(name)
    o = ruby('worksheets',"#{ruby_name.downcase}.rb")
    d = output(name,'defaults')
    o.puts "# coding: utf-8"
    o.puts "# #{name}"
    o.puts
    o.puts "require_relative '../#{compiled_module_name.downcase}'"
    o.puts
    o.puts "module #{compiled_module_name}"
    o.puts "class #{ruby_name.capitalize} < Spreadsheet"
    c.rewrite(i,w,o,d)
    o.puts ""
    close(d)
    if settable_refs
      o.puts "  def initialize"
      d = input(name,'defaults')
      d.lines do |line|
        o.puts line
      end
      o.puts "  end"
      o.puts ""
      close(d)
    end
    o.puts "end"
    o.puts "end"
    close(i,o)
  end

  def compile_worksheet_test(name,xml_filename)
    i = input(name,"values_pruned2.ast")
    ruby_name = ruby_name_for_worksheet_name(name)
    o = ruby('tests',"test_#{ruby_name.downcase}.rb")
    o.puts "# coding: utf-8"
    o.puts "# Test for #{name}"
    o.puts  "require 'test/unit'"
    o.puts  "require_relative '../#{compiled_module_name.downcase}'"
    o.puts
    o.puts "module #{compiled_module_name}"
    o.puts "class Test#{ruby_name.capitalize} < Test::Unit::TestCase"
    o.puts "  def spreadsheet; $spreadsheet ||= Spreadsheet.new; end"
    o.puts "  def worksheet; @worksheet ||= spreadsheet.#{ruby_name}; end"
    CompileToRubyUnitTest.rewrite(i, o)
    o.puts "end"
    o.puts "end"
    close(i,o)
  end
  
  def ruby_name_for_worksheet_name(name)
    unless @worksheet_names
      w = input("worksheet_ruby_names")
      @worksheet_names = Hash[w.readlines.map { |line| line.split("\t").map { |a| a.strip }}]
      close(w)
    end
    @worksheet_names[name]
  end

  
end