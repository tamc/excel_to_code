require_relative '../src/excel_to_code'
this_directory = File.dirname(__FILE__)

command = ExcelToRuby.new

command.excel_file = File.join(this_directory, 'global.xlsx')
command.output_directory = this_directory
command.output_name = 'global'

command.cells_that_can_be_set_at_runtime = { "G.30 (data)" => (4.upto(93).to_a.map { |r| "I#{r}" }) }

command.cells_to_keep = {
  "G.30" => :all,
}

command.actually_compile_code = true
command.actually_run_tests = true

command.run_in_memory = true

command.go!
