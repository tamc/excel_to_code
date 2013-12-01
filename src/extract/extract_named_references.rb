require 'nokogiri'

class ExtractNamedReferences < Nokogiri::XML::SAX::Document 

  attr_accessor :parsing, :input, :output

  def self.extract(input)
    self.new.extract(input)
  end

  def initialize
    super
    @sheet_names = []
  end

  def extract(input)
    @input, @output = input, {}
    @sheet = nil
    @name = nil
    @reference = nil
    parser = Nokogiri::XML::SAX::Parser.new(self)
    parser.parse(input)
    output
  end

  def start_element(name,attributes)
    if name == "sheet"
      @sheet_names << attributes.assoc('name').last
    elsif name == "definedName"
      @sheet = attributes.assoc('localSheetId') && @sheet_names[attributes.assoc('localSheetId').last.to_i]
      @name =  attributes.assoc('name').last
      @reference = ""
    end
  end

  def end_element(name)
    return unless name == "definedName"
    if @sheet
      @output[[@sheet, @name]] = @reference
    else
      @output[@name] = @reference
    end
    @sheet = nil
    @name = nil
    @reference = nil
  end

  def characters(string)
    return unless @reference
    @reference << string
  end

end
