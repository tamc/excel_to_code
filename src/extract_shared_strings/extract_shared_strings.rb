require 'nokogiri'

class ExtractSharedStrings < Nokogiri::XML::SAX::Document 

  attr_accessor :parsing, :input, :output
  
  def self.extract(input,output)
    self.new.extract(input,output)
  end
  
  def initialize
  end
  
  def extract(input,output)
    @input, @output = input, output
    parsing = false
    parser = Nokogiri::XML::SAX::Parser.new(self)
    parser.parse(input)
    output
  end
  
  def start_element(name,attributes)
    self.parsing = true if name == "si"
  end
  
  def end_element(name)
    return unless name == "si"
    self.parsing = false
    output.putc "\n"
  end
  
  def characters(string)
    return unless parsing
    output.write string.gsub("\n","")
  end
  
end