require 'rspec'
require 'stringio'
require 'tmpdir'

# Allow both old and new syntax
RSpec.configure do |config|
  config.example_status_persistence_file_path = File.join(File.dirname(__FILE__), "results.txt")
  config.run_all_when_everything_filtered = true

  config.expect_with :rspec do |c|
    c.syntax = [:should, :expect]
  end
end

require_relative '../src/excel_to_code'

def excel_fragment(name)
  File.open(File.join(File.dirname(__FILE__),'test_data',name))
end

alias :test_data :excel_fragment

class StringIO
  def to_ary
    lines.to_a.map { |l| l.split("\t") }
  end
end

class FunctionTest
  extend ExcelFunctions
  def FunctionTest.original_excel_filename; "filename not specified"; end
end
  
