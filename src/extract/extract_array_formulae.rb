require_relative 'simple_extract_from_xml'

class ExtractArrayFormulae < SimpleExtractFromXML
  
  def start_element(name,attributes)
    if name == 'c'
      @ref = attributes.assoc('r').last
    elsif name == "f" 
      type = attributes.assoc('t')
      if type && type.last == 'array' && attributes.assoc('ref')
        @parsing = true    
        output.write @ref
        output.write "\t"
        output.write attributes.assoc('ref').last
        output.write "\t"
      end
    end
  end
  
  def end_element(name)
    return unless parsing && name == "f"
    self.parsing = false
    output.putc "\n"
  end
  
  def characters(string)
    return unless parsing
    output.write string
  end
  
  
end
