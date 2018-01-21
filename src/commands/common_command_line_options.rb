class CommonCommandLineOptions
  def self.set(command:, options:)
    options.separator ''
    options.separator 'Specific options:'

    options.on('-v', '--version', 'Prints the version number of this code') do
      puts ExcelToCode.version
      exit
    end

    options.on('-o', '--output-name NAME', 'Filename to give to c version of code (and associated ruby interface). Defaults to a folder with the same name as the excel file.') do |name|
      command.output_name = name
    end

    options.on('-r', '--run-tests', 'Compile the generated code and then run the tests') do
      command.actually_run_tests = true
    end

    options.on('-n', '--named-references', 'Transfer named references from spreadsheet to generated code') do
      command.named_references_that_can_be_set_at_runtime = :where_possible
      command.named_references_to_keep = :all
    end

    options.on('-s', '--settable WORKSHEET', 'Make it possible to set the values of cells in this worksheet at runtime. By default no values are settable.') do |sheet|
      command.cells_that_can_be_set_at_runtime = { sheet => :all }
    end

    options.on('-p', '--prune-except WORKSHEET', 'Remove all cells except those on this worksheet, or that are required to calculate values on that worksheet. By default keeps all cells.') do |sheet|
      command.cells_to_keep = { sheet => :all }
    end

    options.on('--isolate WORKSHEET', 'Only performs translation and optimiation of that one worksheet. Useful for debugging an incorrect translation of a large worksheet') do |sheet|
      command.isolate = sheet
    end

    options.on('-d', '--debug', 'Does not perform final optimisations of spreadsheet, leaving the resulting code more similar to the original worksheet, but potentially slower') do |_sheet|
      command.should_inline_formulae_that_are_only_used_once = false
      command.extract_repeated_parts_of_formulae = false
    end

    options.on('-h', '--help', 'Show this message') do
      puts options
      exit
    end
  end
end
