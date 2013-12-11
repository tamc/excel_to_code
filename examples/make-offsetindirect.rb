require_relative '../src/excel_to_code'
this_directory = File.dirname(__FILE__)
command = ExcelToC.new
command.excel_file = File.join(this_directory, 'offsetindirect.xlsx')
command.output_directory = this_directory
command.output_name = 'offsetindirect'
# Handy command:
# cut -f 2 electricity-build-rate-constraint/intermediate/Named\ references\ 000 | pbcopy
command.cells_that_can_be_set_at_runtime = {'Sheet1' => ['A1']}
command.cells_to_keep = {'Sheet1' => :all}
command.actually_compile_code = true
command.actually_run_tests = true
command.run_in_memory = true
command.go!
