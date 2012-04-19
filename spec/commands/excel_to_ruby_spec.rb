require_relative '../spec_helper'

describe ExcelToRuby do
  
  it "Should transform ExampleSpreadsheet.xlsx into the desired ruby code" do
    excel = File.join(File.dirname(__FILE__),'..','test_data','ExampleSpreadsheet.xlsx')
    expected = File.join(File.dirname(__FILE__),'excel_to_ruby_output_expected')
    actual = Dir.mktmpdir
    puts "Writing to #{actual}"
    command = ExcelToRuby.new
    command.excel_file = excel
    command.xml_directory = File.join(actual,'xml')
    command.intermediate_directory = File.join(actual,'intermediate')
    command.output_directory = File.join(actual,'ruby')
    command.output_name = "ExampleSpreadsheet"
    command.cells_that_can_be_set_at_runtime = {
      'Referencing' => ['A4']
    }
    command.go!
    differences = `diff -r #{expected} #{actual}`
    unless differences == ""
      puts
      puts "Differences"
      puts differences
    end
    differences.should == ""
  end
end