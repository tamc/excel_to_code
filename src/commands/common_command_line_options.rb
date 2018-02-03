class CommonCommandLineOptions
  def self.set(command:, options:, generates:, extension:)
    banner = <<-EOT
      #{$0} [options] <excel_file> [output_directory]

      Roughly translate some Excel files into plain #{generates}

      http://github.com/tamc/excel_to_code

      Options:
    EOT
    options.banner = banner.split("\n").map(&:strip).join("\n")

    options.on_tail('-v', '--version', '') do
      puts ExcelToCode.version
      exit
    end

    options.on('-o', '--output-name NAME', "Name the generated code files. Default excelspreadsheet.#{extension}") do |name|
      command.output_name = name
    end

    options.on('-r', '--run-tests', "Test whether the #{generates} matches the Excel.") do
      command.actually_run_tests = true
    end

    options.on('-n', '--named-references', "Include Excel named references as variables in the #{generates}.") do
      command.named_references_that_can_be_set_at_runtime = :where_possible
      command.named_references_to_keep = :all
    end

    options.on('-s', '--settable INPUT_WORKSHEET', "Translate value cells in INPUT_WORKSHEET as settable variables in the #{generates}.") do |sheet|
       
      command.cells_that_can_be_set_at_runtime = { sheet => :all }
    end

    options.on('-p', '--prune-except OUTPUT_WORKSHEET', 'Only translate OUTPUT_WORKSHEET and the cells its results depend on.') do |sheet|
      command.cells_to_keep = { sheet => :all }
    end

    options.on('--isolate FAULTY_WORKSHEET', 'Only translate FAULTY_WORKSHEET. Useful for debugging.') do |sheet|
      command.isolate = sheet
    end

    options.on('-d', '--debug', "Fewer optimisations; the #{generates} should be more similar to the original Excel.") do
      command.should_inline_formulae_that_are_only_used_once = false
      command.extract_repeated_parts_of_formulae = false
    end

    options.on_tail('-h', '--help', '') do
      puts options
      exit
    end
    
    options.set_summary_width 35

  end
  
  def self.parse(options:, command:, arguments:)
    begin
      options.parse!(arguments)
    rescue OptionParser::ParseError => e 
      STDERR.puts e.message, "\n", options 
      return false
    end

    unless arguments.size > 0
      puts options
      return false
    end

    command.excel_file = arguments[0]
    command.output_directory = arguments[1] if arguments[1]
    return true
  end
end
