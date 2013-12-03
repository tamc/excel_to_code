require 'nokogiri'


class ExtractWorksheetTableRelationships < Nokogiri::XML::SAX::Document 

    def self.extract(*args)
      self.new.extract(*args)
    end

    def extract(input)
      @output = [] 
      @parsing = false
      Nokogiri::XML::SAX::Parser.new(self).parse(input)
      @output
    end

  def start_element(name,attributes)
    return false unless name == "tablePart"
    @output << "#{attributes.assoc('r:id').last}"
  end

end
