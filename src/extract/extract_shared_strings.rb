require_relative 'simple_extract_from_xml'

class ExtractSharedStrings < SimpleExtractFromXML

  def start_element(name,attributes)
    self.parsing = true if name == "si"
  end
  
  def end_element(name)
    return unless name == "si"
    self.parsing = false
    output.putc "\n"
  end
  
  def characters(string)
    return unless parsing
    output.write string.gsub("\n","")
  end
  
end