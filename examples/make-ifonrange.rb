require_relative '../src/excel_to_code'
this_directory = File.dirname(__FILE__)
command = ExcelToC.new
command.excel_file = File.join(this_directory,'ifonrange.xlsx')
command.output_directory = this_directory
command.output_name = 'ifonrange'
command.actually_compile_code = true
command.actually_run_tests = true
command.named_references_that_can_be_set_at_runtime = ['input']
command.named_references_to_keep = ['input', 'output'] 
command.go!
