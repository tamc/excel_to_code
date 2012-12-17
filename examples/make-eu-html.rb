require_relative '../src/excel_to_code'
this_directory = File.dirname(__FILE__)
command = ExcelToHTML.new
command.excel_file = File.join(this_directory,'eu.xlsx')
command.output_directory = File.join(this_directory,'eu-html')
command.run_in_memory = false
command.go!
