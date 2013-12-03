require 'nokogiri'

class SimpleExtractFromXML < Nokogiri::XML::SAX::Document 

    def self.extract(*args)
      self.new.extract(*args)
    end

    def extract(sheet_name, input)
      @sheet_name = sheet_name
      @output = {} 
      @parsing = false
      Nokogiri::XML::SAX::Parser.new(self).parse(input)
      @output
    end
    
end
