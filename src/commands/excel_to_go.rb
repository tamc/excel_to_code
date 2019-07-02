# frozen_string_literal: true

require_relative 'excel_to_x'
require 'pathname'

# ExcelToGo turns a spreadsheet into code in the Go language
class ExcelToGo < ExcelToX
  def language
    'go'
  end

  # Skip this
  def replace_values_with_constants; end

  # These actually create the code version of the excel
  def write_code
    write_out_excel_as_code
    write_out_test_as_code
  end

  def write_out_excel_as_code
    log.info 'Starting to write out code'

    f = "#{output_name.downcase}.go"
    o = output(f)

    o.puts "// Compiled version of #{excel_file}"
    o.puts "package #{output_name.downcase}"
    o.puts
    o.puts excel_lib_imports
    o.puts

    c = CompileToGo.new
    c.settable = settable
    c.gettable = gettable
    c.rewrite @formulae, @worksheet_c_names, o
    o.puts

    o.puts excel_lib_functions
    o.puts

    close(o)
    log.info 'Finished writing code'

    format_code(f)
  end

  def write_out_test_as_code
    log.info 'Starting to write out test'

    f = "#{output_name.downcase}_test.go"
    o = output(f)

    o.puts "// Test of compiled version of #{excel_file}"
    o.puts "package #{output_name.downcase}"
    o.puts
    o.puts 'import ('
    o.puts '    "testing"'
    o.puts ')'
    o.puts

    c = CompileToGoTest.new
    c.settable = settable
    c.gettable = gettable
    c.rewrite @formulae, @values, @worksheet_c_names, o
    o.puts

    close(o)
    log.info 'Finished writing tests'

    format_code(f)
  end

  def compile_code
    # Not needed
  end

  def run_tests
    return unless actually_run_tests

    log.info 'Running the resulting tests'
    log.info `cd #{File.join(output_directory)}; go test`
  end

  def format_code(filename)
    log.info 'Running gofmt'
    log.info `gofmt -w -s #{output_path(filename)}`
  end

  def excel_lib
    @excel_lib ||= IO.readlines(path_to_excel_go).join
  end

  def excel_lib_imports
    excel_lib[/import \(.*?\)/m]
  end

  def excel_lib_functions
    excel_lib[/import \(.*?\)(.*)/m, 1]
  end

  def path_to_excel_go
    File.join(File.dirname(__FILE__), '..', 'compile', 'go', 'excel.go')
  end
end
