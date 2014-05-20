require 'ox'

class ExtractNamedReferences < ::Ox::Sax

  attr_accessor :named_references
  attr_accessor :state

  def self.extract(input)
    self.new.extract(input)
  end

  def extract(input_xml)
    @state = :not_parsing
    @sheet_names = [] # This keeps track of the sheet names
    @named_references = {}
    @localSheetId = nil
    @name = nil
    @reference = []
    Ox.sax_parse(self, input_xml, :convert_special => true)
    @named_references
  end

  def start_element(name)
    case name
    when :sheet
      @state = :parsing_sheet_name
    when :definedName
      @state = :parsing_named_reference
    end
  end


  def attr(name, value)
    case state
    when :parsing_sheet_name
      @sheet_names << value if name == :name
    when :parsing_named_reference
      @localSheetId = value.to_i if name == :localSheetId
      @name = value.downcase.to_sym if name == :name
    end
  end

  def end_element(name)
    case name
    when :sheet
      @state = :not_parsing
    when :definedName
      @state = :not_parsing

      reference = @reference.join.gsub('$','')

      @named_references[key] = reference

      @localSheetId = nil
      @name = nil
      @reference = []
    end
  end

  def text(text)
    return unless state == :parsing_named_reference
    @reference << text
  end


  def key
    return @name unless @localSheetId
    sheet = @sheet_names[@localSheetId].downcase.to_sym
    [sheet, @name]
  end

end
