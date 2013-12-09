require_relative '../src/excel_to_code'
this_directory = File.dirname(__FILE__)
command = ExcelToC.new
#command = ExcelToRuby.new
command.excel_file = File.join(this_directory,'model.xlsx')
command.output_directory = this_directory
command.output_name = 'model'
command.cells_that_can_be_set_at_runtime = { "Control" => (5.upto(57).to_a.map { |r| "e#{r}" }) }

command.cells_to_keep = {
  # The names, limits, 10 worders, long descriptions
  "Control" => (5.upto(57).to_a.map { |r| ["d#{r}","f#{r}","h#{r}","i#{r}","j#{r}","k#{r}","bo#{r}","bp#{r}","bq#{r}","br#{r}"] }).flatten, 
  "Intermediate output" => :all, 
  "CostPerCapita" => :all, 
  "Land Use" => :all, 
  "Flows" => :all, 
  "AQ Outputs" => :all, 
}

command.actually_compile_code = true
command.actually_run_tests = true
command.run_in_memory = true
command.go!
