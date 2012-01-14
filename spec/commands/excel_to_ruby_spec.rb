require_relative '../spec_helper'

describe ExcelToRuby do
  
  it "Should transform ExampleSpreadsheet.xlsx into the desired ruby code" do
    excel = File.join(File.dirname(__FILE__),'..','test_data','ExampleSpreadsheet.xlsx')
    expected = File.join(File.dirname(__FILE__),'excel_to_ruby_output_expected')
    actual = Dir.mktmpdir
    puts "Writing to #{actual}"
    command = ExcelToRuby.new
    command.excel_file = excel
    command.output_directory = actual
    command.compiled_module_name = "ExampleSpreadsheet"
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