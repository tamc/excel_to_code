require '../src/excel_to_code'
require 'ruby-prof'

profile =  RubyProf.profile do
  this_directory = File.dirname(__FILE__)
  command = ExcelToC.new
  command.excel_file = File.join(this_directory, 'offsets.xlsx')
  command.output_directory = this_directory
  command.output_name = 'offsets'
  # Handy command:
  # cut -f 2 electricity-build-rate-constraint/intermediate/Named\ references\ 000 | pbcopy
  command.named_references_to_keep = :all
  command.named_references_that_can_be_set_at_runtime = :where_possible
  command.cells_that_can_be_set_at_runtime = :named_references_only
  command.actually_compile_code = true
  command.actually_run_tests = true
  command.run_in_memory = true
  command.go!
end

printer = RubyProf::GraphHtmlPrinter.new(profile)
File.open('profile.html', 'w') do |f|
  printer.print(f)
end

`open profile.html`

