require 'nokogiri'

class ExtractWorksheetDimensions < Nokogiri::XML::SAX::Document 

    def self.extract(*args)
      self.new.extract(*args)
    end

    def extract(input)
      parser = Nokogiri::XML::SAX::Parser.new(self)
      parser.parse(input)
      @output
    end
  
  # FIXME: Is there an elegant way to abort once we have found the dimension tag?
  def start_element(name,attributes)
    return false unless name == "dimension"
    @output = attributes.assoc('ref').last
  end
  
end
