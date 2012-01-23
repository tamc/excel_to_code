require_relative 'simple_extract_from_xml'

class ExtractTable < SimpleExtractFromXML
  
  attr_accessor :worksheet_name
  
  def initialize(worksheet_name = nil)
    super
    @worksheet_name = worksheet_name
  end
  
  def start_element(name,attributes)
    if name == "table"
      output.write "#{attributes.assoc('displayName').last}\t#{@worksheet_name}\t#{attributes.assoc('ref').last}\t#{attributes.assoc('totalsRowCount').try(:last) || 0}"
    elsif name == "tableColumn"
      output.write "\t#{attributes.assoc('name').last}"
    end
  end
  
  def end_element(name)
    return unless name == "table"
    output.putc "\n"
  end  
end
