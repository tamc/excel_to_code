require_relative '../src/excel_to_code'
this_directory = File.dirname(__FILE__)
command = ExcelToC.new
command.excel_file = File.join(this_directory,'getsetranges.xlsx')
command.output_directory = this_directory
command.output_name = 'getsetranges'
command.actually_compile_code = true
command.actually_run_tests = true
command.cells_to_keep = { "Sheet1" => :all }
command.named_references_that_can_be_set_at_runtime = :where_possible # ['A']
command.named_references_to_keep = ['rangeoverblanks'] #['Total', "Not a named reference"]
command.go!
