require_relative '../spec_helper'

describe ExcelToGo do
  it 'Should transform ExampleSpreadsheet.xlsx into the desired ruby code' do
    excel = File.join(File.dirname(__FILE__), '..', 'test_data', 'GoTestSpreadsheet.xlsx')
    actual = File.join(File.dirname(__FILE__), 'excel_to_X_output_actual')
    puts "Writing to #{actual}"
    command = ExcelToGo.new
    command.excel_file = excel
    command.output_directory = File.join(actual, 'go')
    command.output_name = 'examplespreadsheet'
    command.go!
    source_file = File.join(actual, 'go', 'examplespreadsheet.go')
    test_file = File.join(actual, 'go', 'examplespreadsheet_test.go')
    expect(system("go test #{source_file} \"#{test_file}\"")).to be true
  end
end
