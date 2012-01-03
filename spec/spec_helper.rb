require 'rspec'
require 'stringio'
require 'tmpdir'

require_relative '../src/util'
require_relative '../src/commands'
require_relative '../src/excel'
require_relative '../src/extract'
require_relative '../src/rewrite'

def excel_fragment(name)
  File.open(File.join(File.dirname(__FILE__),'test_data',name))
end

alias :test_data :excel_fragment