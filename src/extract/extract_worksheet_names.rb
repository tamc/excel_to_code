require 'ox'

class ExtractWorksheetNames < ::Ox::Sax

    attr_accessor :output
  
    def self.extract(*args)
      self.new.extract(*args)
    end

    def extract(input_xml)
      @output = {}
      @state = :not_parsing
      Ox.sax_parse(self, input_xml, :convert_special => true)
      @output
    end
  
  def start_element(name)
    return false unless name == :sheet
    @state = :parsing
  end

  def attr(attr_name, value)
    return unless @state == :parsing
    case attr_name
    when :name
      @name = value
    when :"r:id"
      @rid = value
    end
  end

  def end_element(name)
    return false unless name == :sheet
    output[@name] = @rid
    @state = :not_parsing
  end
  
end
