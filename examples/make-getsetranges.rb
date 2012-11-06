require_relative '../src/excel_to_code'
this_directory = File.dirname(__FILE__)
command = ExcelToC.new
command.excel_file = File.join(this_directory,'getsetranges.xlsx')
command.output_directory = this_directory
command.output_name = 'getsetranges'
command.actually_compile_code = true
command.actually_run_tests = true
command.run_in_memory = true
command.go!
