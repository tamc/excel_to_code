require 'ox'

class ExtractRelationships < ::Ox::Sax

    attr_accessor :input, :output
  
    def self.extract(*args)
      self.new.extract(*args)
    end

    def extract(input_xml)
      @input, @output = input, {}
      @state = :not_in_relationship
      Ox.sax_parse(self, input_xml, :convert_special => true)
      output
    end
  
  def start_element(name)
    return false unless name == :Relationship
    @state = :in_relationship
  end

  def attr(name, value)
    return unless @state == :in_relationship
    case name
    when :Id
      @id = value
    when :Target
      @target = value
    end
  end 

  def end_element(name)
    return unless @state == :in_relationship
    return unless name == :Relationship
    @output[@id] = @target
    @state = :not_in_relationship
    @id, @target = nil, nil
  end
  
end
