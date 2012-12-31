require_relative '../src/excel_to_code'
this_directory = File.dirname(__FILE__)
command = ExcelToHTML.new
command.excel_file = File.join(this_directory,'2050Model.xlsx')
command.output_directory = File.join(this_directory,'2050Model-html')
command.run_in_memory = true
command.go!
