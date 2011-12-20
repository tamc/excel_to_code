require 'nokogiri'

class SimpleExtractFromXML < Nokogiri::XML::SAX::Document 

    attr_accessor :parsing, :input, :output
  
    def self.extract(input,output)
      self.new.extract(input,output)
    end

    def extract(input,output)
      @input, @output = input, output
      parsing = false
      parser = Nokogiri::XML::SAX::Parser.new(self)
      parser.parse(input)
      output
    end
    
end