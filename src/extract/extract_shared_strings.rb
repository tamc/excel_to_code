require 'nokogiri'

class ExtractSharedStrings < Nokogiri::XML::SAX::Document 

    def self.extract(input)
      self.new.extract(input)
    end

    def extract(input)
      @input, @output = input, []
      @current = nil
      parser = Nokogiri::XML::SAX::Parser.new(self)
      parser.parse(input)
      @output
    end

  def start_element(name,attributes)
    @current = [] if name == "si"
  end
  
  def end_element(name)
    return unless name == "si"
    @output << @current.join
    @current = nil
  end
  
  def characters(string)
    return unless @current
    # FIXME: SHOULDN'T ELINMATE NEWLINES
    @current << string.gsub("\n","")
  end
  
end
