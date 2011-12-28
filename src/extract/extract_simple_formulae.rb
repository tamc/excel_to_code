require_relative 'simple_extract_from_xml'

class ExtractSimpleFormulae < SimpleExtractFromXML
  
  def start_element(name,attributes)
    if name == 'c'
      @ref = attributes.assoc('r').last
    elsif name == "f" && !(attributes.assoc('t'))
      @parsing = true    
      output.write @ref
      output.write "\t"
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
