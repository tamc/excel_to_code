require_relative 'simple_extract_from_xml'

class ExtractValues < SimpleExtractFromXML
  
  attr_accessor :ref, :type
  
  def start_element(name,attributes)
    if name == "v"
      self.parsing = true 
      output.write "#{@ref}\t#{@type}\t"
    elsif name == "c"
      @ref = attributes.assoc('r').last
      type = attributes.assoc('t')
      @type = type ? type.last : "n"
    end
  end
  
  def end_element(name)
    return unless name == "v"
    self.parsing = false
    output.putc "\n"
  end
  
  def characters(string)
    return unless parsing
    output.write string
  end
  
end
