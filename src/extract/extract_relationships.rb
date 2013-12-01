require 'nokogiri'

class ExtractRelationships < Nokogiri::XML::SAX::Document 

    attr_accessor :input, :output
  
    def self.extract(*args)
      self.new.extract(*args)
    end

    def extract(input)
      @input, @output = input, {}
      parser = Nokogiri::XML::SAX::Parser.new(self)
      parser.parse(input)
      output
    end
  
  def start_element(name,attributes)
    return false unless name == "Relationship"
    @output[attributes.assoc('Id').last] = attributes.assoc('Target').last
  end
  
end
