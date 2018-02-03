require_relative '../../src/excel_to_code'
this_directory = File.dirname(__FILE__)
command = ExcelToC.new
command.excel_file = File.join(this_directory,'example.xlsx')
command.output_directory = this_directory
command.output_name = 'excelspreadsheet'
command.actually_compile_code = true
command.actually_run_tests = true
command.named_references_that_can_be_set_at_runtime = ['in'] 
command.named_references_to_keep = ['out']
command.go!
