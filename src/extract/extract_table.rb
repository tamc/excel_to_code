require 'nokogiri'

class ExtractTable < Nokogiri::XML::SAX::Document 

    def self.extract(*args)
      self.new.extract(*args)
    end

    def extract(sheet_name, input)
      @sheet_name = sheet_name
      @output = {} 
      @table_name = nil
      @table_ref = nil
      @table_total_rows = nil
      @table_columns = nil

      @parsing = false
      Nokogiri::XML::SAX::Parser.new(self).parse(input)
      @output
    end
    
  def start_element(name,attributes)
    if name == "table"
      @table_name = attributes.assoc('displayName').last
      @table_ref = attributes.assoc('ref').last
      @table_total_rows = attributes.assoc('totalsRowCount').try(:last) || "0"
      @table_columns = []
    elsif name == "tableColumn"
      @table_columns << attributes.assoc('name').last
    end
  end
  
  def end_element(name)
    return unless name == "table"
    @output[@table_name] = [@sheet_name, @table_ref, @table_total_rows, *@table_columns]
  end  
end
