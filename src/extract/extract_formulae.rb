require 'nokogiri'

class ExtractFormulae < Nokogiri::XML::SAX::Document 

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

  def start_element(name,attributes)
    if name == 'c'
      @ref = attributes.assoc('r').last
    elsif name == "f"
      type = attributes.assoc('t')
      @formula = []
      start_formula( type && type.last, attributes)
    end
  end

  def end_element(name)
    return unless @parsing && name == "f"
    @parsing = false
    write_formula
  end

  def characters(string)
    return unless @parsing
    @formula.push(string)
  end

  def start_formula(type,attributes)
    # Override
  end

  def write_formula   
    # Override
  end

end
