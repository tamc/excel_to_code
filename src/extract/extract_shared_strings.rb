require 'ox'

class ExtractSharedStrings < ::Ox::Sax

    def self.extract(input)
      self.new.extract(input)
    end

    def extract(input_xml)
      @output = []
      @current = nil
      Ox.sax_parse(self, input_xml, :convert_special => true)
      @output
    end

  def start_element(name)
    @current = [] if name == :si
  end
  
  def end_element(name)
    return unless name == :si
    @output << [:string, @current.join]
    @current = nil
  end
  
  def text(string)
    return unless @current
    # FIXME: SHOULDN'T ELINMATE NEWLINES
    @current << string#.gsub("\n","")
  end
  
end
