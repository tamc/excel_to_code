require_relative 'simple_extract_from_xml'

class ExtractWorksheetTableRelationships < SimpleExtractFromXML

  def start_element(name,attributes)
    return false unless name == "tablePart"
    output.puts "#{attributes.assoc('r:id').last}"
  end

end