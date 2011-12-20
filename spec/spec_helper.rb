require 'rspec'

def excel_fragment(name)
  File.open(File.join(File.dirname(__FILE__),'excel_fragments',name))
end