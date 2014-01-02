require_relative '../src/excel_to_code'
this_directory = File.dirname(__FILE__)
command = ExcelToC.new
# command = ExcelToRuby.new
command.excel_file = File.join(this_directory,'utf8-strings.xlsx')
command.output_directory = this_directory
command.output_name = 'utf8strings'
command.actually_compile_code = true
command.actually_run_tests = true
command.cells_that_can_be_set_at_runtime = { '2015' => ['B3'] }
command.go!
