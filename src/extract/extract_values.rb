require_relative 'simple_extract_from_xml'

class ExtractValues < SimpleExtractFromXML

  def start_element(name,attributes)
    self.parsing = true if name == "v"
    return unless name == "c"
    output.write attributes.assoc('r').last
    output.write "\t"
    type = attributes.assoc('t')
    if type
      output.write type.last
    else
      output.write "n"
    end
    output.write "\t"
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
