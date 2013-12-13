require_relative '../src/excel_to_code'
this_directory = File.dirname(__FILE__)

command = ExcelToC.new

command.excel_file = File.join(this_directory, 'global.xlsx')
command.output_directory = 'ext'
command.output_name = 'global'

command.cells_that_can_be_set_at_runtime = { "User inputs" => (7.upto(46).to_a.map { |r| "E#{r}" }) }

command.cells_to_keep = {
  "User inputs" => :all,
  "Detailed lever guids" => :all,
  "Outputs - Climate impacts" => :all,
  "Outputs - Emissions" => :all,
  "Outputs - Energy" => :all,
  "Outputs - Land use, technology" => :all,
  "Outputs - Costs" => :all,
  "Outputs - Energy flows" => :all,
}

command.actually_compile_code = true
command.actually_run_tests = true

command.run_in_memory = true

command.go!
