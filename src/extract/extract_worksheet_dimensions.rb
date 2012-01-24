require_relative 'simple_extract_from_xml'

class ExtractWorksheetDimensions < SimpleExtractFromXML
  
  # FIXME: Is there an elegant way to abort once we have found the dimension tag?
  def start_element(name,attributes)
    return false unless name == "dimension"
    output.puts "#{attributes.assoc('ref').last}"
  end
  
end
