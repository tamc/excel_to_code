require_relative 'simple_extract_from_xml'

class ExtractRelationships < SimpleExtractFromXML
  
  def start_element(name,attributes)
    return false unless name == "Relationship"
    output.puts "#{attributes.assoc('Id').last}\t#{attributes.assoc('Target').last}"
  end
  
end
