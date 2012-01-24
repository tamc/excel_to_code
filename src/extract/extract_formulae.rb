require_relative 'simple_extract_from_xml'

class ExtractFormulae < SimpleExtractFromXML
  
  attr_accessor :ref, :formula
  
  def start_element(name,attributes)
    if name == 'c'
      @ref = attributes.assoc('r').last
    elsif name == "f"
      type = attributes.assoc('t')
      @formula = []
      start_formula( type && type.last, attributes)
    end
  end
  
  def start_formula(type,attributes)
    # Should be overriden in sub classes
  end
  
  def write_formula
    # Should be overriden in sub classes
  end
  
  def end_element(name)
    return unless parsing && name == "f"
    self.parsing = false
    write_formula
  end
  
  def characters(string)
    return unless parsing
    @formula.push(string)
  end
  
end
