require_relative '../src/excel_to_code'
this_directory = File.dirname(__FILE__)
command = ExcelToC.new
#command = ExcelToRuby.new
command.excel_file = File.join(this_directory,'eu.xlsx')
command.output_directory = this_directory
command.output_name = 'eu'
command.actually_compile_code = true
command.actually_run_tests = true
command.write_tests_in_c = true
command.go!
