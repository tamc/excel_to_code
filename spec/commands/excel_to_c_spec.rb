require_relative '../spec_helper'

describe ExcelToC do
  
  it "Should transform ExampleSpreadsheet.xlsx into the desired c code" do
    excel = File.join(File.dirname(__FILE__),'..','test_data','ExampleSpreadsheet.xlsx')
    expected = File.join(File.dirname(__FILE__),'excel_to_X_output_expected')
    actual = File.join(File.dirname(__FILE__),'excel_to_X_output_actual')
    puts "Writing to #{actual}"
    command = ExcelToC.new
    command.excel_file = excel
    command.output_directory = File.join(actual,'c')
    command.output_name = "ExampleSpreadsheet"
    #command.cells_that_can_be_set_at_runtime = {
    #  'Referencing' => ['A4']
    #}
    command.actually_compile_code = true
    command.go!
    test_file = File.join(actual,'c','test_examplespreadsheet.rb')
    expect(system("ruby \"#{test_file}\"")).to be true

  end
end
