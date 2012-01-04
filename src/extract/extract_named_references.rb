require_relative 'simple_extract_from_xml'

class ExtractNamedReferences < SimpleExtractFromXML
  
  attr_accessor :sheet_names
  
  def initialize
    super
    @sheet_names = {}
  end
  
  def start_element(name,attributes)
    if name == "sheet"
      sheet_names[attributes.assoc('sheetId').last.strip] = attributes.assoc('name').last
    elsif name == "definedName"
      @parsing = true    
      sheet = attributes.assoc('localSheetId')
      if sheet
        output.write sheet_names[sheet.last.strip]
      end
      output.write "\t"
      output.write attributes.assoc('name').last
      output.write "\t"
    end
  end
  
  def end_element(name)
    return unless name == "definedName"
    self.parsing = false
    output.putc "\n"
  end
  
  def characters(string)
    return unless parsing
    output.write string
  end
  
end
