require_relative '../spec_helper'

describe ExcelToRuby do
  
  it "Should transform ExampleSpreadsheet.xlsx into the desired ruby code" do
    excel = File.join(File.dirname(__FILE__),'..','test_data','ExampleSpreadsheet.xlsx')
    expected = File.join(File.dirname(__FILE__),'excel_to_X_output_expected')
    actual = File.join(File.dirname(__FILE__),'excel_to_X_output_actual')
    puts "Writing to #{actual}"
    command = ExcelToRuby.new
    command.excel_file = excel
    command.xml_directory = File.join(actual,'xml')
    command.intermediate_directory = File.join(actual,'intermediate')
    command.output_directory = File.join(actual,'ruby')
    command.output_name = "RubyExampleSpreadsheet"
    #command.cells_that_can_be_set_at_runtime = {
    #  'Referencing' => ['A4']
    #}
    command.run_in_memory = true
    command.go!
    require_relative File.join(actual,'ruby','test_rubyexamplespreadsheet')
    Minitest.run.should == true
  end
end
