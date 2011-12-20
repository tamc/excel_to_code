require_relative 'simple_extract_from_xml'

class ExtractWorksheetNames < SimpleExtractFromXML
  
  def start_element(name,attributes)
    return false unless name == "sheet"
    output.puts "#{attributes.assoc('r:id').last}\t#{attributes.assoc('name').last}"
  end
  
end
