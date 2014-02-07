require_relative '../src/excel_to_code'
this_directory = File.dirname(__FILE__)
command = ExcelToC.new
#command = ExcelToRuby.new
command.excel_file = File.join(this_directory,'rangearrayformulaesa.xlsx')
command.output_directory = this_directory
command.output_name = 'rangearrayformulaesa'
command.cells_that_can_be_set_at_runtime = { "Growth Paths " => :all }
command.cells_to_keep = {
    "Growth Paths " => :all,
    "IND.a" => :all, 
  }

command.extract_repeated_parts_of_formulae = false

command.actually_compile_code = true
command.actually_run_tests = true
command.go!
