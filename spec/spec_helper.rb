require 'rspec'
require 'stringio'
require 'tmpdir'

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
  
