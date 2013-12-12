require_relative '../src/excel_to_code'
this_directory = File.dirname(__FILE__)
command = ExcelToC.new
command.excel_file = File.join(this_directory,'arrayformulatest.xlsx')
command.output_directory = this_directory
command.output_name = 'arrayformulatest'
command.actually_compile_code = true
command.actually_run_tests = true
command.named_references_that_can_be_set_at_runtime = :where_possible # ['A']
command.named_references_to_keep = :all # ['Total']
command.run_in_memory = true
command.go!
